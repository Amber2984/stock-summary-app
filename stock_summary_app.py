import streamlit as st
import pandas as pd
import io

def process_file(uploaded_file):
    # 读取 Excel 文件
    excel_file = pd.ExcelFile(uploaded_file)
    df = excel_file.parse('明细')

    # 删除多余列
    if 'Unnamed: 14' in df.columns:
        df = df.drop(columns=['Unnamed: 14'])

    # 重命名列
    df.columns = [
        '客户代码', '客户姓名', '证券代码', '证券名称', '证券类别', '业务标示',
        '成交金额', '手续费', '买卖方向', '交收日期', '服务人员', '部门',
        '是否签约', '双融账户'
    ]

    # 类型转换 & 新增辅助列
    df['成交金额'] = df['成交金额'].astype(float)
    df['手续费'] = df['手续费'].astype(float)
    df['是否签约客户'] = df['是否签约'].notna() & (df['是否签约'] != '#N/A')
    df['是否双融账户'] = df['双融账户'].notna()

    # 汇总表
    summary = df.groupby(['交收日期', '证券名称']).apply(
        lambda x: pd.Series({
            '买入客户数': x[x['买卖方向'] == '证券买入']['客户代码'].nunique(),
            '总成交金额（万）': round(x[x['买卖方向'] == '证券买入']['成交金额'].sum() / 10000, 2),
            '总佣金收入（元）': round(x['手续费'].sum(), 2),

            '其中签约客户数': x[x['是否签约客户']]['客户代码'].nunique(),
            '其中签约成交金额（万）': round(x[x['是否签约客户']]['成交金额'].sum() / 10000, 2),
            '签约佣金收入（元）': round(x[x['是否签约客户']]['手续费'].sum(), 2),
            '签约客户佣金占比': round(
                x[x['是否签约客户']]['手续费'].sum() / x['手续费'].sum(), 4) if x['手续费'].sum() > 0 else 0,

            '双融账户买入户数': x[(x['是否双融账户']) & (x['买卖方向'] == '证券买入')]['客户代码'].nunique(),
            '双融账户买入金额（万）': round(x[(x['是否双融账户']) & (x['买卖方向'] == '证券买入')]['成交金额'].sum() / 10000, 2),
            '双融账户佣金收入（元）': round(x[x['是否双融账户']]['手续费'].sum(), 2),
        })
    ).reset_index()

    return summary

# Streamlit 页面布局
st.title("📊 签约客户股票交易数据统计工具")
st.markdown("上传你的股票明细Excel文件（需包含名为 ‘明细’ 的sheet），系统将自动生成汇总统计表。")

uploaded_file = st.file_uploader("请上传Excel文件", type=["xlsx"])

if uploaded_file is not None:
    try:
        result_df = process_file(uploaded_file)
        st.success("✅ 处理成功！以下是统计结果：")
        st.dataframe(result_df)

        # 提供下载按钮
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            result_df.to_excel(writer, index=False, sheet_name="汇总结果")
        st.download_button(
            label="📥 下载Excel汇总结果",
            data=output.getvalue(),
            file_name="股票交易统计汇总.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    except Exception as e:
        st.error(f"❌ 处理失败: {e}")
