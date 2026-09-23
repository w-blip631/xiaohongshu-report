import streamlit as st
import pandas as pd
import datetime
import io

st.set_page_config(page_title="小红书月报生成器", page_icon="📊", layout="wide")

st.title("📊 小红书月报自动生成器")
st.caption("上传官方导出的笔记明细表，一键生成月度报表")

# ================= 上传文件 =================
uploaded_file = st.file_uploader("上传小红书导出的「笔记列表明细表.xlsx」", type=["xlsx", "xls"])

if uploaded_file is not None:
   # 读取数据（跳过第一行标题，第二行是表头）
    df = pd.read_excel(uploaded_file, skiprows=1)
    df.columns = df.columns.str.strip()  # 清理列名
    df = df.loc[:, ~df.columns.str.contains('^Unnamed')]  # 删掉多余的空列
    
    # 清理列名（去除空格）
    df.columns = df.columns.str.strip()
    
    # 转换日期格式（中文日期 → 标准日期）
    def parse_chinese_date(s):
        try:
            s = str(s)
            s = s.replace('年', '-').replace('月', '-').replace('日', ' ')
            s = s.replace('时', ':').replace('分', ':').replace('秒', '')
            return pd.to_datetime(s)
        except:
            return pd.NaT
    
    df['发布时间'] = df['首次发布时间'].apply(parse_chinese_date)
    df['周'] = df['发布时间'].dt.isocalendar().week
    df['互动总量'] = df['点赞'] + df['评论'] + df['收藏']
    df['互动率'] = (df['互动总量'] / df['观看量'] * 100).round(2)
    
    # ================= 手动输入外部数据 =================
    st.sidebar.header("✏️ 手动填写外部数据")
    tmall_visitors = st.sidebar.number_input("天猫总进店人数", min_value=0, value=0, step=1)
    brand_search = st.sidebar.number_input("品牌词进店人数", min_value=0, value=0, step=1)
    paid_search = st.sidebar.number_input("拍立淘进店人数", min_value=0, value=0, step=1)
    
    # ================= 核心数据概览 =================
    st.subheader("一、核心数据概览")
    
    total_exposure = int(df['曝光'].sum())
    total_views = int(df['观看量'].sum())
    total_likes = int(df['点赞'].sum())
    total_comments = int(df['评论'].sum())
    total_collects = int(df['收藏'].sum())
    total_follows = int(df['涨粉'].sum())
    total_shares = int(df['分享'].sum())
    total_interactions = total_likes + total_comments + total_collects
    avg_interaction_rate = round(total_interactions / total_views * 100, 2) if total_views > 0 else 0
    avg_click_rate = round(df['封面点击率'].mean() * 100, 2)
    
    # 小红书引流占比
    xhs_visitors = brand_search + paid_search
    xhs_ratio = round(xhs_visitors / tmall_visitors * 100, 2) if tmall_visitors > 0 else 0
    
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("总曝光", f"{total_exposure:,}")
    col2.metric("总观看", f"{total_views:,}")
    col3.metric("总互动量", f"{total_interactions:,}")
    col4.metric("总涨粉", f"{total_follows:,}")
    
    col5, col6, col7, col8 = st.columns(4)
    col5.metric("平均互动率", f"{avg_interaction_rate}%")
    col6.metric("平均点击率", f"{avg_click_rate}%")
    col7.metric("小红书引流占比", f"{xhs_ratio}%")
    col8.metric("笔记总数", f"{len(df)}篇")
    
    # ================= 单篇笔记明细 =================
    st.subheader("二、单篇笔记明细")
    
    detail_df = df[['笔记标题', '发布时间', '体裁', '曝光', '观看量', 
                     '点赞', '评论', '收藏', '涨粉', '分享', '互动总量', '互动率']].copy()
    detail_df = detail_df.sort_values('发布时间', ascending=False)
    
    st.dataframe(detail_df, use_container_width=True, hide_index=True)
    
    # ================= 周度对比 =================
    st.subheader("三、周度数据对比")
    
    weekly = df.groupby('周').agg({
        '曝光': 'sum',
        '观看量': 'sum',
        '点赞': 'sum',
        '评论': 'sum',
        '收藏': 'sum',
        '涨粉': 'sum',
        '互动总量': 'sum',
        '笔记标题': 'count'
    }).rename(columns={'笔记标题': '笔记数'}).reset_index()
    
    weekly['平均互动率'] = (weekly['互动总量'] / weekly['观看量'] * 100).round(2)
    
    st.dataframe(weekly, use_container_width=True, hide_index=True)
    
    # ================= 导出 Excel =================
    st.subheader("四、下载报表")
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # 写入核心数据
        summary_data = {
            '指标': ['总曝光', '总观看', '总点赞', '总评论', '总收藏', '总涨粉', 
                     '总分享', '总互动量', '平均互动率(%)', '平均点击率(%)', 
                     '天猫总进店人数', '小红书引流人数', '小红书引流占比(%)', '笔记总数'],
            '数据': [total_exposure, total_views, total_likes, total_comments, 
                     total_collects, total_follows, total_shares, total_interactions,
                     avg_interaction_rate, avg_click_rate, tmall_visitors, 
                     xhs_visitors, xhs_ratio, len(df)]
        }
        summary_df = pd.DataFrame(summary_data)
        summary_df.to_excel(writer, sheet_name='核心数据概览', index=False)
        
        # 写入单篇明细
        detail_df.to_excel(writer, sheet_name='单篇笔记明细', index=False)
        
        # 写入周度对比
        weekly.to_excel(writer, sheet_name='周度对比', index=False)
        
        # 给表头加样式
        workbook = writer.book
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': '#A8C0D6',
            'font_color': 'white',
            'border': 1
        })
        
        for sheet_name in ['核心数据概览', '单篇笔记明细', '周度对比']:
            worksheet = writer.sheets[sheet_name]
            for col_num, value in enumerate(
                summary_df.columns if sheet_name == '核心数据概览' 
                else (detail_df.columns if sheet_name == '单篇笔记明细' else weekly.columns)
            ):
                worksheet.write(0, col_num, value, header_format)
                worksheet.set_column(col_num, col_num, 18)
    
    output.seek(0)
    st.download_button(
        label="📥 下载完整 Excel 报表",
        data=output,
        file_name=f"小红书月报_{datetime.datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
