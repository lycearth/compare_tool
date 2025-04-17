import pandas as pd
import re
# --- 配置区 ---
DENY_KEYWORDS = {
    "平纹", "人字纹", "斜纹", "缎纹",
    "加厚", "加宽", "不抓毛", "射出勾", "勾面", "毛面", "自粘"
}

# --- 分词与匹配判断工具 ---
def normalize_token_list(s):
    s = re.sub(r'[^\u4e00-\u9fa5a-zA-Z0-9]', ' ', str(s))
    return sorted(s.lower().split())

def safe_fuzzy_match(name1, name2, deny_sensitive=True):
    if not name1 or not name2:
        return False

    tokens1 = normalize_token_list(name1)
    tokens2 = normalize_token_list(name2)

    if not tokens1 or not tokens2:
        return False

    if tokens1 == tokens2:
        return True

    common = set(tokens1) & set(tokens2)
    if not common:
        return False

    if deny_sensitive:
        for kw in DENY_KEYWORDS:
            in1 = kw in name1
            in2 = kw in name2
            if in1 and in2:
                continue
            if in1 or in2:
                return False

    return len(common) >= 2 and abs(len(tokens1) - len(tokens2)) <= 1

import pandas as pd

def build_final_table(df_matched, df_unmatched_p, df_unmatched_q):
    rows = []

    def safe_float(x):
        try:
            return round(float(x), 4)
        except:
            return None

    # ✅ 匹配成功的数据
    for _, row in df_matched.iterrows():
        diff_cons = safe_float(row["采购_单耗"]) - safe_float(row["报价_单耗"]) \
            if pd.notna(row["采购_单耗"]) and pd.notna(row["报价_单耗"]) else None
        diff_price = safe_float(row["采购_单价"]) - safe_float(row["报价_单价"]) \
            if pd.notna(row["采购_单价"]) and pd.notna(row["报价_单价"]) else None

        rows.append({
            "采购标识": row["采购_标识"],
            "报价标识": row["报价_标识"],
            "采购单耗": row["采购_单耗"],
            "报价单耗": row["报价_单耗"],
            "单耗差值": diff_cons,
            "采购单价": row["采购_单价"],
            "报价单价": row["报价_单价"],
            "单价差值": diff_price,
            "匹配方式": row.get("匹配方式", "匹配成功"),
            "匹配状态": "✅"
        })

    # ✅ 仅采购未匹配
    for _, row in df_unmatched_p.iterrows():
        rows.append({
            "采购标识": row["采购_标识"],
            "报价标识": None,
            "采购单耗": row["采购_单耗"],
            "报价单耗": None,
            "单耗差值": None,
            "采购单价": row["采购_单价"],
            "报价单价": None,
            "单价差值": None,
            "匹配方式": "未匹配（仅采购）",
            "匹配状态": "❌"
        })

    # ✅ 仅报价未匹配
    for _, row in df_unmatched_q.iterrows():
        rows.append({
            "采购标识": None,
            "报价标识": row["报价_标识"],
            "采购单耗": None,
            "报价单耗": row["报价_单耗"],
            "单耗差值": None,
            "采购单价": None,
            "报价单价": row["报价_单价"],
            "单价差值": None,
            "匹配方式": "未匹配（仅报价）",
            "匹配状态": "❌"
        })

    df = pd.DataFrame(rows)
    df.dropna(how="all", inplace=True)  # 清除全空行
    return df


def highlight_diff(val):
    try:
        val = float(val)
        if val > 0:
            return "background-color: #FFECEC"  # 浅红
        elif val < 0:
            return "background-color: #E8F5E9"  # 浅绿
    except:
        return ""
    return ""
