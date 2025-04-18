import streamlit as st
import pandas as pd
from io import BytesIO
from comparison_engine import compare_data
from utils import build_final_table, highlight_diff, normalize_token_list, safe_fuzzy_match

# 页面配置
st.set_page_config(page_title="采购与报价比对工具", layout="wide")
st.title("📦 采购与报价比对工具")
st.markdown(
    "上传采购与报价 Excel 或 CSV 文件，选择表头所在行与比对列，即可进行产品标识（名称/编号）、单耗、单价的自动比对。"
)

# 上传文件
col1, col2 = st.columns(2)
with col1:
    purchase_file = st.file_uploader("📥 上传采购文件", type=["xlsx","xls","csv"], key="p_file")
with col2:
    quote_file    = st.file_uploader("📥 上传报价文件", type=["xlsx","xls","csv"], key="q_file")

# 辅助函数：根据别名选索引

def smart_index(options, *aliases):
    for alias in aliases:
        if alias in options:
            return options.index(alias)
    return 0

# 辅助：安全提取标识

def safe_identity(val, fallback):
    if pd.isna(val) or not str(val).strip():
        return fallback
    return str(val).strip()

# 渲染匹配下拉框

def render_matching_column(item_list, container):
    for idx, row in item_list:
        pid = safe_identity(row["采购_标识"], f"行{idx}")
        if idx not in st.session_state.manual_matches:
            # 自动推荐逻辑
            best, best_score = None, 0
            toks_p = set(normalize_token_list(pid))
            for _, rq in st.session_state.df_unmatched_q.iterrows():
                qid = safe_identity(rq["报价_标识"], "")
                common = toks_p & set(normalize_token_list(qid))
                if qid and safe_fuzzy_match(pid, qid) and len(common) > best_score:
                    best_score, best = len(common), qid
            if best:
                st.session_state.manual_matches[idx] = best
        # 构造下拉选项
        opts = ["（不匹配）"] + [safe_identity(rq["报价_标识"], "未知") for _, rq in st.session_state.df_unmatched_q.iterrows()]
        default = st.session_state.manual_matches.get(idx, "（不匹配）")
        sel = container.selectbox(f"为采购项【{pid}】选报价：", opts, index=opts.index(default), key=f"sel_{idx}")
        if sel != "（不匹配）":
            st.session_state.manual_matches[idx] = sel
        else:
            st.session_state.manual_matches.pop(idx, None)

# 应用人工匹配逻辑

def apply_manual_matches():
    df_unmatched_p = st.session_state["df_unmatched_p"]
    df_unmatched_q = st.session_state["df_unmatched_q"]
    df_matched     = st.session_state["df_matched"]
    manual_matches = st.session_state["manual_matches"]
    applied = 0
    for p_idx, q_ident in list(manual_matches.items()):
        rows = df_unmatched_q[df_unmatched_q["报价_标识"] == q_ident]
        if not rows.empty and p_idx in df_unmatched_p.index:
            rq = rows.iloc[0]
            new = {
                "采购_标识": df_unmatched_p.at[p_idx, "采购_标识"],
                "报价_标识": rq["报价_标识"],
                "采购_单耗": df_unmatched_p.at[p_idx, "采购_单耗"],
                "报价_单耗": rq["报价_单耗"],
                "采购_单价": df_unmatched_p.at[p_idx, "采购_单价"],
                "报价_单价": rq["报价_单价"],
                "匹配方式": "人工匹配"
            }
            st.session_state.df_matched     = pd.concat([df_matched, pd.DataFrame([new])], ignore_index=True)
            st.session_state.df_unmatched_q = df_unmatched_q[df_unmatched_q["报价_标识"] != q_ident]
            st.session_state.df_unmatched_p = df_unmatched_p.drop(p_idx)
            applied += 1
    st.session_state.manual_matches = {}
    st.success(f"✅ 共应用 {applied} 条人工匹配")

# 主流程：上传后选择表头、映射、自动比对
if purchase_file and quote_file:
    st.subheader("👀 表头行选择")
    purch_preview = pd.read_excel(purchase_file, header=None, nrows=10)
    quote_preview = pd.read_excel(quote_file, header=None, nrows=10)

    st.write("📑 采购文件前10行")
    st.dataframe(purch_preview)
    purch_header = st.selectbox("选择采购文件表头行 (从0开始)", range(10), index=0, key="p_header")

    st.write("📑 报价文件前10行")
    st.dataframe(quote_preview)
    quote_header = st.selectbox("选择报价文件表头行 (从0开始)", range(10), index=0, key="q_header")

    df_purch = pd.read_excel(purchase_file, header=purch_header)
    df_quote = pd.read_excel(quote_file, header=quote_header)

    st.subheader("🔧 手动列映射")
    purch_cols = ["无"] + df_purch.columns.astype(str).tolist()
    quote_cols = ["无"] + df_quote.columns.astype(str).tolist()

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 📦 采购字段映射")
        purch_identity = st.selectbox("产品标识列（采购）", purch_cols)
        purch_price    = st.selectbox("单价列（采购）", purch_cols, index=smart_index(purch_cols, "单价", "price"))
        purch_cons     = st.selectbox("单耗列（采购）", purch_cols, index=smart_index(purch_cols, "单耗", "consumption"))
    with col2:
        st.markdown("#### 📋 报价字段映射")
        quote_identity = st.selectbox("产品标识列（报价）", quote_cols)
        quote_price    = st.selectbox("单价列（报价）", quote_cols, index=smart_index(quote_cols, "单价", "price"))
        quote_cons     = st.selectbox("单耗列（报价）", quote_cols, index=smart_index(quote_cols, "单耗", "consumption"))

    if st.button("🚀 开始比对"):
        df_matched, df_unmatched_p, df_unmatched_q = compare_data(
            df_purch, df_quote,
            {"identity": purch_identity, "price": purch_price, "consumption": purch_cons},
            {"identity": quote_identity, "price": quote_price, "consumption": quote_cons}
        )
        st.session_state.df_matched     = df_matched
        st.session_state.df_unmatched_p = df_unmatched_p
        st.session_state.df_unmatched_q = df_unmatched_q
        st.session_state.manual_matches = {}
        st.success("✅ 自动比对完成，请继续人工比对或导出结果")

if "df_unmatched_p" in st.session_state and "df_unmatched_q" in st.session_state:
    # 初始化展开状态：默认折叠
    if "expander_open" not in st.session_state:
        st.session_state.expander_open = False

    # 检查是否已有选择行为，则自动展开
    for p_idx in st.session_state.df_unmatched_p.index:
        if f"sel_{p_idx}" in st.session_state:
            st.session_state.expander_open = True
            break

    # 折叠面板：实时渲染所有下拉选择框
    with st.expander("🔎 未匹配 - 人工指定报价项", expanded=st.session_state.expander_open):
        c1, c2 = st.columns(2)
        items = list(st.session_state.df_unmatched_p.iterrows())
        mid = len(items) // 2

        # 左侧一半
        for idx, row in items[:mid]:
            pid = safe_identity(row["采购_标识"], f"行{idx}")
            opts = ["（不匹配）"] + [
                safe_identity(rq["报价_标识"], "未知")
                for _, rq in st.session_state.df_unmatched_q.iterrows()
            ]
            sel = c1.selectbox(f"为采购项【{pid}】选报价：", opts, key=f"sel_{idx}")

        # 右侧一半
        for idx, row in items[mid:]:
            pid = safe_identity(row["采购_标识"], f"行{idx}")
            opts = ["（不匹配）"] + [
                safe_identity(rq["报价_标识"], "未知")
                for _, rq in st.session_state.df_unmatched_q.iterrows()
            ]
            sel = c2.selectbox(f"为采购项【{pid}】选报价：", opts, key=f"sel_{idx}")

    # 普通按钮一次性应用所有已选映射
    if st.button("✅ 应用人工匹配并更新结果表"):
        # 1) 收集所有 selectbox 当前状态
        manual_matches = {}
        for p_idx in st.session_state.df_unmatched_p.index:
            key = f"sel_{p_idx}"
            if key in st.session_state and st.session_state[key] != "（不匹配）":
                manual_matches[p_idx] = st.session_state[key]

        # 2) 批量应用
        applied = 0
        for p_idx, q_ident in manual_matches.items():
            df_p = st.session_state.df_unmatched_p
            df_q = st.session_state.df_unmatched_q
            rows = df_q[df_q["报价_标识"] == q_ident]
            if not rows.empty:
                rq = rows.iloc[0]
                new = {
                    "采购_标识": df_p.at[p_idx, "采购_标识"],
                    "报价_标识": rq["报价_标识"],
                    "采购_单耗": df_p.at[p_idx, "采购_单耗"],
                    "报价_单耗": rq["报价_单耗"],
                    "采购_单价": df_p.at[p_idx, "采购_单价"],
                    "报价_单价": rq["报价_单价"],
                    "匹配方式": "人工匹配"
                }
                st.session_state.df_matched = pd.concat(
                    [st.session_state.df_matched, pd.DataFrame([new])],
                    ignore_index=True
                )
                # 从未匹配列表中移除已匹配行
                st.session_state.df_unmatched_q = df_q[df_q["报价_标识"] != q_ident]
                st.session_state.df_unmatched_p = df_p.drop(p_idx)
                applied += 1

        # 3) 清除已应用行对应的 selectbox 状态，其余保留
        for p_idx in manual_matches:
            sel_key = f"sel_{p_idx}"
            if sel_key in st.session_state:
                del st.session_state[sel_key]

        st.success(f"✅ 共应用 {applied} 条人工匹配")

    # 最后——显示并导出综合结果
    final_df = build_final_table(
        st.session_state.df_matched,
        st.session_state.df_unmatched_p,
        st.session_state.df_unmatched_q
    )
    styled = final_df.style.applymap(
        highlight_diff, subset=["单耗差值","单价差值"]
    ).format(precision=2, na_rep="")

    st.subheader("📊 综合比对结果总表")
    st.dataframe(styled, use_container_width=True)

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        final_df.to_excel(writer, index=False, sheet_name="比对结果总表")
        wb = writer.book
        ws = writer.sheets["比对结果总表"]
        red_fmt   = wb.add_format({'bg_color':'#FFCCCC'})
        green_fmt = wb.add_format({'bg_color':'#C8E6C9'})
        for col in ["单耗差值","单价差值"]:
            if col in final_df.columns:
                idx = final_df.columns.get_loc(col)
                for i, v in enumerate(final_df[col]):
                    try:
                        vv = float(v)
                        fmt = red_fmt if vv>0 else green_fmt if vv<0 else None
                        if fmt:
                            ws.write(i+1, idx, vv, fmt)
                    except:
                        pass
    buffer.seek(0)
    st.download_button(
        "📥 下载比对结果 Excel（含颜色）",
        data=buffer.getvalue(),
        file_name="对比结果_总表_含高亮.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

