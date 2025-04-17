import pandas as pd
import re
from utils import normalize_token_list, safe_fuzzy_match
# --- 主比对逻辑 ---
def compare_data(purchase_input, quote_input, purch_mapping, quote_mapping) -> tuple:
    import pandas as pd
    import re
    from comparison_engine import normalize_token_list, safe_fuzzy_match

    std_cols = {
        'identity': 'product_identity',  # ✅ 用户映射时只选这个字段即可
        'price': 'price',
        'consumption': 'consumption'
    }

    def load_and_rename(df_or_file, mapping):
        if not isinstance(df_or_file, pd.DataFrame):
            try:
                df = pd.read_excel(df_or_file)
            except:
                df = pd.read_csv(df_or_file)
        else:
            df = df_or_file.copy()

        rename_map = {v: std_cols[k] for k, v in mapping.items() if v and v != "无"}
        df.rename(columns=rename_map, inplace=True)
        for std in std_cols.values():
            if std not in df.columns:
                df[std] = None
        return df

    # ✅ 支持中文 + 去符号 + 忽略大小写
    def normalize_identity(x):
        return re.sub(r"[^\u4e00-\u9fa5a-zA-Z0-9]", "", str(x)).upper().strip()

    # --- 加载数据 ---
    p = load_and_rename(purchase_input, purch_mapping)
    q = load_and_rename(quote_input, quote_mapping)

    # --- 清洗标识字段 ---
    for df in [p, q]:
        df["product_identity"] = df["product_identity"].fillna("").astype(str).str.strip()
        df["product_identity_std"] = df["product_identity"].apply(normalize_identity)

    for col in ['price', 'consumption']:
        p[col] = pd.to_numeric(p[col], errors='coerce')
        q[col] = pd.to_numeric(q[col], errors='coerce')

    p = p.where(pd.notnull(p), None)
    q = q.where(pd.notnull(q), None)

    matched_pairs = []
    unmatched_p = set(p.index)
    unmatched_q = set(q.index)

    print("✔ 采购字段：", p.columns.tolist())
    print(p.head(3)[["product_identity", "price", "consumption"]])

    # ✅ 精确匹配 - 使用标准化标识字段
    code_lookup = {q.at[j, 'product_identity_std']: j for j in unmatched_q if q.at[j, 'product_identity_std']}
    for i in list(unmatched_p):
        code = p.at[i, 'product_identity_std']
        if code and code in code_lookup:
            j = code_lookup[code]
            matched_pairs.append((i, j, "精确匹配 - 标识"))
            unmatched_p.discard(i)
            unmatched_q.discard(j)

    # ✅ 模糊匹配 - 剩余未匹配项
    for i in list(unmatched_p):
        identity_p = p.at[i, 'product_identity_std']
        if not identity_p.strip():
            continue

        best_j = None
        best_score = -1
        tokens_p = set(normalize_token_list(identity_p))

        for j in unmatched_q:
            identity_q = q.at[j, 'product_identity_std']
            if not identity_q.strip():
                continue

            tokens_q = set(normalize_token_list(identity_q))
            common = tokens_p & tokens_q
            score = len(common)

            if not safe_fuzzy_match(identity_p, identity_q, deny_sensitive=True):
                continue

            if score > best_score:
                best_score = score
                best_j = j

        if best_j is not None:
            matched_pairs.append((i, best_j, "模糊匹配 - 标识"))
            unmatched_p.discard(i)
            unmatched_q.discard(best_j)

    # ✅ 构建比对结果
    matched_rows = []
    unmatched_p_rows = []
    unmatched_q_rows = []

    for i, j, method in matched_pairs:
        row_p = p.loc[i]
        row_q = q.loc[j]
        matched_rows.append({
            "采购_标识": row_p['product_identity'],
            "报价_标识": row_q['product_identity'],
            "采购_单耗": row_p['consumption'],
            "报价_单耗": row_q['consumption'],
            "采购_单价": row_p['price'],
            "报价_单价": row_q['price'],
            "匹配方式": method
        })

    for i in unmatched_p:
        row = p.loc[i]
        unmatched_p_rows.append({
            "采购_标识": row['product_identity'],
            "采购_单耗": row['consumption'],
            "采购_单价": row['price']
        })

    for j in unmatched_q:
        row = q.loc[j]
        unmatched_q_rows.append({
            "报价_标识": row['product_identity'],
            "报价_单耗": row['consumption'],
            "报价_单价": row['price']
        })

    df_matched = pd.DataFrame(matched_rows)
    df_unmatched_p = pd.DataFrame(unmatched_p_rows)
    df_unmatched_q = pd.DataFrame(unmatched_q_rows)

    for df in [df_matched, df_unmatched_p, df_unmatched_q]:
        df.replace({None: pd.NA, "": pd.NA}, inplace=True)
        df.dropna(how="all", inplace=True)

    return df_matched, df_unmatched_p, df_unmatched_q




