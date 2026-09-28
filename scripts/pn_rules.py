# -*- coding: utf-8 -*-
"""
PN 編碼規則對照表，資料來源：90PN_Rule.xlsx（Region & Adapter / Channel / 保留碼 / 外包商代碼 / EMS Org 等分頁）。
規則變動時，只需要改這個檔案，不用動 generate_pn.py 的邏輯。
"""

# SKU -> (2.4G Channel, Region 代碼(第11碼), Adapter 代碼(第12碼), Power Adapter 文字(Description用))
# 註：APAC 在來源表裡依 2.4G Channel 不同分兩列，這裡拆成 APAC_11 / APAC_13 兩個獨立選項。
REGION_ADAPTER = {
    "Dummy":   {"ch": None, "region": "0", "adapter": "0", "adapter_label": "Dummy / WW"},
    "US":      {"ch": 11, "region": "A", "adapter": "1", "adapter_label": "P_US"},
    "MX":      {"ch": 11, "region": "B", "adapter": "1", "adapter_label": "P_US"},
    "CN":      {"ch": 13, "region": "C", "adapter": "5", "adapter_label": "P_CN"},
    "APAC_11": {"ch": 11, "region": "E", "adapter": "3", "adapter_label": "P_EU", "note": "2024.09起不使用"},
    "APAC_AU": {"ch": 11, "region": "F", "adapter": "A", "adapter_label": "P_EU_UK_AU"},
    "APAC_13": {"ch": 13, "region": "G", "adapter": "A", "adapter_label": "P_EU_UK_AU", "note": "2024.09起不使用"},
    "HK":      {"ch": 11, "region": "H", "adapter": "2", "adapter_label": "P_UK"},
    "IN":      {"ch": 11, "region": "I", "adapter": "C", "adapter_label": "P_IN"},
    "JP":      {"ch": 13, "region": "J", "adapter": "6", "adapter_label": "P_JP"},
    "KR":      {"ch": 13, "region": "K", "adapter": "4", "adapter_label": "P_KR"},
    "IL":      {"ch": 13, "region": "L", "adapter": "3", "adapter_label": "P_EU"},
    "WEU":     {"ch": 13, "region": "M", "adapter": "3", "adapter_label": "P_EU"},
    "EEU":     {"ch": 13, "region": "N", "adapter": "3", "adapter_label": "P_EU"},
    "EU":      {"ch": 13, "region": "O", "adapter": "3", "adapter_label": "P_EU"},
    "AR":      {"ch": 11, "region": "P", "adapter": "B", "adapter_label": "P_AR"},
    "RU":      {"ch": 13, "region": "R", "adapter": "3", "adapter_label": "P_EU"},
    "SG":      {"ch": 11, "region": "S", "adapter": "2", "adapter_label": "P_UK"},
    "TW":      {"ch": 11, "region": "T", "adapter": "1", "adapter_label": "P_US"},
    "UK":      {"ch": 13, "region": "U", "adapter": "9", "adapter_label": "P_EU_UK"},
    "WW":      {"ch": None, "region": "W", "adapter": "0", "adapter_label": "Dummy / WW"},
    "CA":      {"ch": 11, "region": "X", "adapter": "1", "adapter_label": "P_US"},
    "BZ":      {"ch": 11, "region": "Y", "adapter": "8", "adapter_label": "P_BZ"},
    "AU":      {"ch": 13, "region": "Z", "adapter": "7", "adapter_label": "P_AU", "note": "2024.09起不使用"},
}

# Channel (90碼第8碼)
CHANNEL_CODE = {
    "0": "Default",
    "A": "Amazon",
    "B": "BBY",
    "W": "Walmart",
}

# Remark（90碼 G欄，Description 用；目前只定義一個值，未來會擴充）
REMARK_CODE = {
    "PW": "預設加密",
}

# 保留碼（90碼第15碼，V欄 = 生產地）
RESERVE_CODE_90 = {
    "0": "中國生產 MIC",
    "9": "福利品 Refurbish",
    "C": "CSC (For 客服入賬使用)",
    "S": "For 參展樣品 (限NPI機種使用)",
    "T": "台灣生產 MIT",
    "V": "越南生產 MIV",
}

# 外包商代碼（90碼第13碼，T欄；單字母，代表整個公司）
VENDOR_CODE = {
    "A": "智易 Arcadyan",
    "B": "訊舟 Acelink",
    "C": "同維/共進",
    "D": "亞旭 Askey",
    "E": "金磚",
    "F": "啟碁",
    "G": "晶訊",
    "H": "世紀本原",
    "J": "華芸",
    "K": "品威",
    "L": "東碩",
    "M": "康全",
    "N": "正文",
    "P": "微網優聯",
    "Q": "微網力合",
    "R": "偉迪",
}

# EMS Org 三碼 -> (公司說明, 對應的外包商代碼單字母)
EMS_ORG = {
    "AGP": ("智易 Arcadyan_大陸", "A"),
    "AJR": ("智易 Arcadyan_台灣", "A"),
    "AR1": ("智易 Arcadyan_越南", "A"),
    "AFL": ("訊舟 Acelink_大陸", "B"),
    "AQR": ("訊舟 Acelink_台灣", "B"),
    "AWM": ("同維/共進_深圳", "C"),
    "AUG": ("同維/共進_香港", "C"),
    "AOR": ("亞旭 Askey_大陸", "D"),
    "AQH": ("亞旭 Askey_台灣", "D"),
    "ATC": ("正文_崑山", "N"),
    "AZK": ("正文_安博", "N"),
    "AS9": ("正文_越南", "N"),
    "AS3": ("世紀本原", "H"),
    "AK4": ("華芸", "J"),
    "AKM": ("啟碁", "F"),
    "AKN": ("金磚", "E"),
    "AFW": ("晶訊", "G"),
    "AVT": ("品威", "K"),
    "AWJ": ("東碩_台灣", "L"),
    "AXY": ("康全_大陸", "M"),
    "A3C": ("偉迪", "R"),
    "AKS": ("訊舟 Acelink_大陸(Consign)", "B"),
    "ASI": ("訊舟 Acelink_台灣(Consign)", "B"),
}

# 90碼第10碼（屬性）：目前流程只支援 M
ATTRIBUTE_90 = {
    "M": "MP (Default)",
}

# ---- 80編碼原則 ----

# ID+Color（80碼第8碼）
ID_COLOR = {
    "0": "不分色",
    "B": "Black",
    "W": "White",
}

# PCBA（80碼第11碼）
PCBA_CODE = {
    "0": "All SKU 通用 (Default)",
    "A": "PCBA_US",
    "B": "PCBA_EU",
    "E": "EPA",
    "I": "IPA",
}

# 機種分類（80碼第12碼）
PRODUCT_CLASS = {
    "0": "不分子母機 (一般 Router/PCE/USB 等)",
    "R": "子母機 - Router (母機)",
    "N": "子母機 - Node (子機)",
}

# 出貨機台數量（80碼第13碼），目前最多到 3
SHIP_QTY = {
    "1": "1pk",
    "2": "2pk",
    "3": "3pk",
}

# 屬性（80碼第10碼）：固定 M
ATTRIBUTE_80 = {
    "M": "Default",
}

# 保留碼（80碼第15碼）：目前固定用 0，之後若要擴充，只需要在這裡加值。
RESERVE_CODE_80 = {
    "0": "Default",
}

PROJECT_TYPES = ["OEM", "ODM", "JDM", "Cosign"]
