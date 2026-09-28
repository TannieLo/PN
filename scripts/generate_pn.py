# -*- coding: utf-8 -*-
"""
讀取 Issue Form 解析出來的 JSON，產生/更新 90PN_<Model Name>.xlsx。

用法：
    python generate_pn.py <issue_data.json> <repo_root> <output_json>

<issue_data.json>：由 GitHub Action 用 issue-forms-body-parser 解析 Issue 內容後輸出的 JSON，
                    key 是 .github/ISSUE_TEMPLATE/pn-request.yml 裡每個欄位的 id。
<repo_root>：repo 根目錄（放 90PN_model name_template.xlsx 的地方）。
<output_json>：本次執行結果（是否成功、錯誤訊息、產生的檔名等），供後續 workflow 步驟讀取。

規則本身（對照表）都放在 pn_rules.py，這裡只處理「怎麼把表單資料套進規則、寫進 Excel」的邏輯。
"""
import json
import os
import sys

import openpyxl

import pn_rules as R

TEMPLATE_NAME = "90PN_model name_template.xlsx"
NO_RESPONSE = "_No response_"

GROUP_COUNT = 5
SKU_COUNT = 10


def norm(value):
    """issue-forms-body-parser 對未填寫的欄位會回填 '_No response_' 或 None，統一轉成空字串。"""
    if value is None:
        return ""
    value = str(value).strip()
    if value == NO_RESPONSE:
        return ""
    return value


def code_of(value):
    """下拉選單選項格式固定為 '<code> - <說明>'，取出 code。若不含 ' - ' 就整段當 code。"""
    value = norm(value)
    if not value:
        return ""
    return value.split(" - ", 1)[0].strip()


def get(data, key):
    return norm(data.get(key, ""))


def build_header(data):
    return {
        "model_name": get(data, "model_name"),
        "project_code": get(data, "project_code"),
        "project_type": code_of(get(data, "project_type")),
        "ems_org": code_of(get(data, "ems_org")),
    }


def build_groups(data):
    groups = {}
    for i in range(1, GROUP_COUNT + 1):
        color = code_of(get(data, f"g{i}_color"))
        pcba = code_of(get(data, f"g{i}_pcba"))
        pclass = code_of(get(data, f"g{i}_class"))
        pk = code_of(get(data, f"g{i}_pk"))
        pn70 = get(data, f"g{i}_70pn")
        desc70 = get(data, f"g{i}_70desc")
        pn60 = get(data, f"g{i}_60pn")
        desc60 = get(data, f"g{i}_60desc")
        if not any([color, pcba, pclass, pk, pn70, desc70, pn60, desc60]):
            continue
        groups[str(i)] = {
            "idx": i,
            "color": color,
            "pcba": pcba,
            "pclass": pclass,
            "pk": pk,
            "pn70": pn70,
            "desc70": desc70,
            "pn60": pn60,
            "desc60": desc60,
        }
    return groups


def build_skus(data):
    skus = []
    for i in range(1, SKU_COUNT + 1):
        sku = code_of(get(data, f"s{i}_sku"))
        if not sku:
            continue
        skus.append({
            "idx": i,
            "sku": sku,
            "channel": code_of(get(data, f"s{i}_channel")) or "0",
            "remark": code_of(get(data, f"s{i}_remark")),
            "reserve": code_of(get(data, f"s{i}_reserve")) or "0",
            "group_ref": code_of(get(data, f"s{i}_group")),
        })
    return skus


def validate(header, groups, skus):
    errors = []

    if not header["model_name"]:
        errors.append("缺少 Model Name。")
    if len(header["project_code"]) != 3:
        errors.append(f"Project Code 必須是 3 個字元，目前填的是「{header['project_code']}」。")
    if header["project_type"] not in R.PROJECT_TYPES:
        errors.append(f"專案類型「{header['project_type']}」不是可辨識的選項（OEM/ODM/JDM/Cosign）。")
    if header["ems_org"] and header["ems_org"] not in R.EMS_ORG:
        errors.append(f"代工廠 EMS「{header['ems_org']}」不在 EMS Org 對照表裡。")

    if not skus:
        errors.append("沒有填寫任何一組 SKU，至少需要一組才能產生檔案。")

    for s in skus:
        if s["sku"] not in R.REGION_ADAPTER:
            errors.append(f"第 {s['idx']} 組 SKU「{s['sku']}」不在 Region & Adapter 對照表裡。")
        if not s["group_ref"]:
            errors.append(f"第 {s['idx']} 組 SKU 沒有指定要用哪一個 80 PN 分組。")
        elif s["group_ref"] not in groups:
            errors.append(f"第 {s['idx']} 組 SKU 指定的 80 PN 分組「{s['group_ref']}」沒有被定義。")

    # 同一批次裡，若兩個分組的 (顏色/PCBA/機種分類/PK) 簽章完全相同，代表其實是同一顆 80 PN，
    # 這種情況下 70/60 PN 也必須一致，否則就是使用者填表時互相矛盾，不能自動決定要用哪一組。
    sig_to_groups = {}
    for gid, g in groups.items():
        sig = (g["color"], g["pcba"], g["pclass"], g["pk"])
        sig_to_groups.setdefault(sig, []).append(gid)
    for sig, gids in sig_to_groups.items():
        if len(gids) < 2:
            continue
        first = groups[gids[0]]
        for other_id in gids[1:]:
            other = groups[other_id]
            if (first["pn70"], first["desc70"], first["pn60"], first["desc60"]) != \
               (other["pn70"], other["desc70"], other["pn60"], other["desc60"]):
                errors.append(
                    f"分組 {gids[0]} 跟分組 {other_id} 的顏色/PCBA/機種分類/出貨機台數量完全相同"
                    f"（代表應該是同一顆 80 PN），但填的 70/60 PN 或品規不一樣，請合併成同一組再重新送出。"
                )

    for gid, g in groups.items():
        if header["project_type"] == "OEM":
            if not (g["pn70"] and g["pn60"]):
                errors.append(f"專案類型是 OEM，但分組 {gid} 沒有填 70 PN / 60 PN。")

    return errors


def signature_90(row):
    """90碼第14碼(流水號)以外的 14 個欄位組成的簽章，用來判斷是不是跟其他列重複。"""
    return (
        row["H"], row["I"], row["J"], row["K"], row["L"], row["M"], row["N"],
        row["O"], row["P"], row["Q"], row["R"], row["S"], row["T"], row["V"],
    )


def signature_80(group):
    return (group["color"], group["pcba"], group["pclass"], group["pk"])


def next_free_row(ws, col, start_row):
    row = start_row
    while ws.cell(row=row, column=col).value not in (None, ""):
        row += 1
    return row


def col_letter_to_index(letter):
    return openpyxl.utils.column_index_from_string(letter)


def read_existing_90(ws):
    """讀出既有 90編碼原則 分頁裡，row5 開始每一列 H~V 的值（含流水號），供 append 時比對。"""
    rows = []
    row = 5
    while ws.cell(row=row, column=col_letter_to_index("C")).value not in (None, ""):
        rows.append({
            letter: ws.cell(row=row, column=col_letter_to_index(letter)).value
            for letter in "HIJKLMNOPQRSTUV"
        })
        row += 1
    return rows


def scan_other_files_same_project_code(repo_root, project_code, exclude_path):
    """
    掃描 repo 裡所有 90PN_*.xlsx（除了正在處理的這個檔案），找出 Project Code 相同的檔案，
    彙整它們既有的 90/80 碼簽章。90 PN / 80 PN 本身不含 Model Name，所以不同 Model 只要共用
    Project Code，流水號就必須放在一起比對，否則會算出重複的 PN。
    """
    sig90 = {}
    sig80 = {}
    exclude_abs = os.path.abspath(exclude_path)
    for fname in os.listdir(repo_root):
        if not fname.startswith("90PN_") or fname == TEMPLATE_NAME or not fname.lower().endswith(".xlsx"):
            continue
        fpath = os.path.join(repo_root, fname)
        if os.path.abspath(fpath) == exclude_abs:
            continue
        try:
            wb = openpyxl.load_workbook(fpath, data_only=True)
        except Exception:
            continue
        if "90編碼原則" not in wb.sheetnames:
            continue
        rows90 = read_existing_90(wb["90編碼原則"])
        if not rows90:
            continue
        file_project_code = "".join(str(rows90[0][k]) for k in "LMN")
        if file_project_code != project_code:
            continue
        for r in rows90:
            sig = tuple(r[k] for k in "HIJKLMNOPQRSTV")
            sig90.setdefault(sig, set()).add(str(r["U"]))
        if "80 編碼原則" in wb.sheetnames:
            for r in read_existing_80(wb["80 編碼原則"]):
                sig = (r["G"], r["H"], r["I"], r["J"])
                sig80[sig] = {
                    "serial": r["K"],
                    "pn70": r["C"], "desc70": r["D"],
                    "pn60": r["E"], "desc60": r["F"],
                    "source_file": fname,
                }
    return sig90, sig80


def validate_cross_file_conflicts(groups, referenced_group_ids, cross_file_sig80):
    """
    80 PN 不含 Model Name，所以如果分組的（顏色/PCBA/機種分類/PK）組合，跟另一個機種檔案裡
    既有的組合一樣，理論上代表兩者共用同一顆 80 PN——但若兩邊填的 70/60 PN 不同，代表這其實是
    兩個不同的硬體，只是規則上湊巧撞碼，不能自動判斷該用哪一邊，要請人工確認後再調整分組內容。
    """
    errors = []
    for gid in referenced_group_ids:
        g = groups[gid]
        info = cross_file_sig80.get(signature_80(g))
        if info is None:
            continue
        if (info["pn70"], info["desc70"], info["pn60"], info["desc60"]) != \
           (g["pn70"], g["desc70"], g["pn60"], g["desc60"]):
            errors.append(
                f"分組 {gid} 的顏色/PCBA/機種分類/出貨機台數量組合，跟另一個機種檔案「{info['source_file']}」"
                f"裡既有的 80 PN 組合相同（該檔案是 70PN={info['pn70'] or '(空)'}／60PN={info['pn60'] or '(空)'}），"
                f"但這次填的 70/60 PN 不一樣。這兩個機種可能被誤判成共用同一顆 80 PN，"
                f"請確認是否真的要共用，或調整分組內容避免撞碼。"
            )
    return errors


def append_90_rows(ws, header, skus, groups, cross_file_sig90):
    existing = read_existing_90(ws)
    used_serials_by_sig = {sig: set(serials) for sig, serials in cross_file_sig90.items()}
    for r in existing:
        sig = tuple(r[k] for k in "HIJKLMNOPQRSTV")
        used_serials_by_sig.setdefault(sig, set()).add(str(r["U"]))

    item_col = col_letter_to_index("A")
    next_item_no = 1
    row_no = 5
    while ws.cell(row=row_no, column=item_col).value not in (None, ""):
        v = ws.cell(row=row_no, column=item_col).value
        try:
            next_item_no = max(next_item_no, int(v) + 1)
        except (TypeError, ValueError):
            pass
        row_no += 1
    write_row = row_no

    project_code = header["project_code"]
    generated = []
    for s in skus:
        ra = R.REGION_ADAPTER[s["sku"]]
        base = {
            "H": 9, "I": 0, "J": "I", "K": "G",
            "L": project_code[0], "M": project_code[1], "N": project_code[2],
            "O": s["channel"],
            "P": "-",
            "Q": "M",
            "R": ra["region"],
            "S": ra["adapter"],
            "T": R.EMS_ORG.get(header["ems_org"], (None, ""))[1],
            "V": s["reserve"],
        }
        sig = tuple(base[k] for k in "HIJKLMNOPQRSTV")
        used = used_serials_by_sig.setdefault(sig, set())
        serial = 0
        while str(serial) in used:
            serial += 1
        used.add(str(serial))

        ws.cell(row=write_row, column=item_col, value=next_item_no)
        ws.cell(row=write_row, column=col_letter_to_index("B"),
                value=f'=C{write_row}&"/"&D{write_row}&"/"&E{write_row}&"/"&F{write_row}&"//"&G{write_row}')
        ws.cell(row=write_row, column=col_letter_to_index("C"), value=header["model_name"])
        ws.cell(row=write_row, column=col_letter_to_index("D"), value=s["sku"])
        ws.cell(row=write_row, column=col_letter_to_index("E"), value=ra["ch"])
        ws.cell(row=write_row, column=col_letter_to_index("F"), value=ra["adapter_label"])
        ws.cell(row=write_row, column=col_letter_to_index("G"), value=s["remark"])
        for letter in "HIJKLMNOPQRST":
            ws.cell(row=write_row, column=col_letter_to_index(letter), value=base[letter])
        ws.cell(row=write_row, column=col_letter_to_index("U"), value=serial)
        ws.cell(row=write_row, column=col_letter_to_index("V"), value=s["reserve"])

        pn90 = "".join(str(base[k]) for k in "HIJKLMN") + str(s["channel"]) + "-" + \
            "M" + str(ra["region"]) + str(ra["adapter"]) + \
            str(R.EMS_ORG.get(header["ems_org"], (None, ""))[1] or "") + str(serial) + str(s["reserve"])
        generated.append({"sku_idx": s["idx"], "sku": s["sku"], "group_ref": s["group_ref"], "pn90": pn90})

        write_row += 1
        next_item_no += 1

    return generated


def read_existing_80(ws):
    rows = []
    row = 12
    while ws.cell(row=row, column=col_letter_to_index("A")).value not in (None, ""):
        rows.append({
            letter: ws.cell(row=row, column=col_letter_to_index(letter)).value
            for letter in "ABCDEFGHIJKL"
        })
        row += 1
    return rows


def append_80_rows(ws, header, groups, referenced_group_ids, cross_file_sig80):
    existing = read_existing_80(ws)
    sig_to_serial = {sig: info["serial"] for sig, info in cross_file_sig80.items()}
    max_serial = -1
    for serial in list(sig_to_serial.values()) + [r["K"] for r in existing]:
        try:
            max_serial = max(max_serial, int(serial))
        except (TypeError, ValueError):
            pass
    for r in existing:
        sig = (r["G"], r["H"], r["I"], r["J"])
        sig_to_serial[sig] = r["K"]

    own_file_sigs = {(r["G"], r["H"], r["I"], r["J"]) for r in existing}

    next_row = next_free_row(ws, col_letter_to_index("A"), 12)
    project_code = header["project_code"]
    reserve = "0"

    result_by_group = {}
    for gid in referenced_group_ids:
        g = groups[gid]
        sig = signature_80(g)

        if sig in own_file_sigs:
            # 這個組合這份檔案裡已經寫過了（例如同一個 Model 之前的申請），不用再新增一列。
            result_by_group[gid] = {"pn80": None, "serial": sig_to_serial[sig], "reused": True}
            continue

        reused_from_other_file = sig in sig_to_serial
        if reused_from_other_file:
            serial = sig_to_serial[sig]
        else:
            max_serial += 1
            serial = max_serial
            sig_to_serial[sig] = serial

        pn80 = f"80IG{project_code}{g['color']}-M{g['pcba']}{g['pclass']}{g['pk']}{serial}{reserve}"

        ws.cell(row=next_row, column=col_letter_to_index("A"), value=gid)
        ws.cell(row=next_row, column=col_letter_to_index("B"), value=pn80)
        ws.cell(row=next_row, column=col_letter_to_index("C"), value=g["pn70"])
        ws.cell(row=next_row, column=col_letter_to_index("D"), value=g["desc70"])
        ws.cell(row=next_row, column=col_letter_to_index("E"), value=g["pn60"])
        ws.cell(row=next_row, column=col_letter_to_index("F"), value=g["desc60"])
        ws.cell(row=next_row, column=col_letter_to_index("G"), value=g["color"])
        ws.cell(row=next_row, column=col_letter_to_index("H"), value=g["pcba"])
        ws.cell(row=next_row, column=col_letter_to_index("I"), value=g["pclass"])
        ws.cell(row=next_row, column=col_letter_to_index("J"), value=g["pk"])
        ws.cell(row=next_row, column=col_letter_to_index("K"), value=serial)
        ws.cell(row=next_row, column=col_letter_to_index("L"), value=reserve)

        result_by_group[gid] = {"pn80": pn80, "serial": serial, "reused": reused_from_other_file}
        next_row += 1

    return result_by_group


def unmerge_from_row(ws, start_row):
    """
    範本裡 80 編碼原則 分頁從 row11 開始有合併儲存格（給範例用的排版），
    我們要用自己的表格排版重新利用這塊區域，所以先拆開，避免寫入 MergedCell 出錯。
    """
    for merged_range in list(ws.merged_cells.ranges):
        if merged_range.min_row >= start_row:
            ws.unmerge_cells(str(merged_range))


def ensure_80_header(ws):
    """把複製過來的範例列清掉、寫上我們自己這套簡化過的表格標題（只在第一次寫入時做）。"""
    marker = ws.cell(row=10, column=col_letter_to_index("A")).value
    if marker == "群組編號":
        return
    headers = ["群組編號", "80 PN", "70 PN", "70 品規", "60 PN", "60 品規",
               "ID+Color", "PCBA", "機種分類", "出貨機台數量", "流水號", "保留碼"]
    for i, text in enumerate(headers):
        ws.cell(row=10, column=col_letter_to_index("A") + i, value=text)


def load_or_create_workbook(repo_root, model_name):
    target_name = f"90PN_{model_name}.xlsx"
    target_path = os.path.join(repo_root, target_name)
    if os.path.exists(target_path):
        wb = openpyxl.load_workbook(target_path)
        is_new = False
    else:
        template_path = os.path.join(repo_root, TEMPLATE_NAME)
        wb = openpyxl.load_workbook(template_path)
        is_new = True
    return wb, target_path, is_new


def main():
    issue_json_path, repo_root, output_json_path = sys.argv[1], sys.argv[2], sys.argv[3]
    with open(issue_json_path, "r", encoding="utf-8") as f:
        data = json.load(f)

    header = build_header(data)
    groups = build_groups(data)
    skus = build_skus(data)

    errors = validate(header, groups, skus)
    result = {"ok": not errors, "errors": errors}

    if errors:
        with open(output_json_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return

    wb, target_path, is_new = load_or_create_workbook(repo_root, header["model_name"])
    referenced_group_ids = sorted({s["group_ref"] for s in skus})

    cross_sig90, cross_sig80 = scan_other_files_same_project_code(
        repo_root, header["project_code"], target_path)

    cross_errors = validate_cross_file_conflicts(groups, referenced_group_ids, cross_sig80)
    if cross_errors:
        result["ok"] = False
        result["errors"] = cross_errors
        with open(output_json_path, "w", encoding="utf-8") as f:
            json.dump(result, f, ensure_ascii=False, indent=2)
        return

    ws90 = wb["90編碼原則"]
    ws80 = wb["80 編碼原則"]

    if is_new:
        # 清掉範本裡的範例列（row 5），從乾淨的狀態開始寫。
        for col in range(1, 24):
            ws90.cell(row=5, column=col).value = None

    unmerge_from_row(ws80, 10)
    ensure_80_header(ws80)
    if is_new:
        for row in range(8, 15):
            for col in range(1, 17):
                ws80.cell(row=row, column=col).value = None
        ensure_80_header(ws80)

    group_results = append_80_rows(ws80, header, groups, referenced_group_ids, cross_sig80)
    sku_results = append_90_rows(ws90, header, skus, groups, cross_sig90)

    os.makedirs(os.path.dirname(target_path) or ".", exist_ok=True)
    wb.save(target_path)

    result.update({
        "file_path": os.path.relpath(target_path, repo_root),
        "model_name": header["model_name"],
        "is_new_file": is_new,
        "sku_count": len(skus),
        "group_count": len(referenced_group_ids),
        "generated_90pn": sku_results,
        "generated_80pn": group_results,
    })
    with open(output_json_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    main()
