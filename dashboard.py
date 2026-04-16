import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 1. CẤU HÌNH TRANG VÀ CSS
st.set_page_config(page_title="Dashboard Năng Suất & Kết Quả Khảo Sát Sale", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    /* Ẩn toàn bộ dấu vết Streamlit Header/Menu */
    header {visibility: hidden !important;}
    #MainMenu {visibility: hidden !important;}
    footer {visibility: hidden !important;}
    
    /* Ép nền trắng tuyệt đối cho Container chính và Sidebar (Light Theme override) */
    [data-testid='stAppViewContainer'] { background-color: white !important; color: black !important; }
    [data-testid='stSidebar'] { background-color: #f8f9fa !important; }
    
    /* Ép màu chữ */
    .stMarkdown, .stText, h1, h2, h3, p { color: #111111 !important; }
    
    /* Thiết kế form card KPI chuyên nghiệp theo yêu cầu */
    div[data-testid="metric-container"] {
        background-color: #ffffff;
        border: 1px solid #e0e0e0;
        border-radius: 10px;
        box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.08);
        padding: 15px;
        margin-bottom: 20px;
    }
    
    /* Tối ưu chữ trong KPI */
    div[data-testid="metric-container"] label {
        font-weight: 600 !important;
        font-size: 16px !important;
        color: #555555;
    }
    
    /* CSS bo góc bảng Insight */
    .insight-box {
        background-color: #f8f9fa;
        border-left: 5px solid #17a2b8;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 20px;
    }
    
    /* Ép định dạng chữ cho bảng HTML tĩnh */
    th {text-align: center !important; font-weight: bold !important; color: #111111 !important; font-size: 14px !important;}
    td {text-align: center !important;}
    </style>
""", unsafe_allow_html=True)


@st.cache_data
def load_and_preprocess_data(file_path):
    df = pd.read_excel(file_path, sheet_name=0, usecols="A:J")
    df.columns = df.columns.astype(str).str.strip()
    
    if 'Inside' in df.columns:
        df['Inside'] = df['Inside'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df['Inside'] = df['Inside'].replace('nan', 'Unknown')
    else:
        df['Inside'] = 'Unknown'
        
    if 'Mail FPT' in df.columns:
        df['Mail FPT'] = df['Mail FPT'].astype(str).replace('nan', '')
        
    if 'Họ và tên đầy đủ' in df.columns:
        df['Họ và tên đầy đủ'] = df['Họ và tên đầy đủ'].astype(str).str.strip()
    
    df['Search_Name'] = df['Inside'] + " - " + df.get('Họ và tên đầy đủ', '')
    
    df['Phòng'] = df.get('Phòng', pd.Series(['Unknown'] * len(df))).fillna('Unknown')
    df['Tháng'] = df.get('Tháng', pd.Series(['Unknown'] * len(df))).astype(str).fillna('Unknown')
    df['Team'] = df.get('Team', pd.Series(['Unknown'] * len(df))).fillna('Unknown')
    df['Leader'] = df.get('Leader', pd.Series(['Unknown'] * len(df))).fillna('Unknown')
    df['Đánh giá'] = df.get('Đánh giá', pd.Series(['Trống'] * len(df))).fillna('Trống')
    
    df['Đánh giá_Clean'] = df['Đánh giá'].str.capitalize()
    
    if 'Họ và tên đầy đủ' in df.columns:
        df = df[~df['Họ và tên đầy đủ'].str.contains('Tổng|Total|Average', case=False, na=False)]
        df = df[df['Họ và tên đầy đủ'] != 'nan']

    return df


FILE_PATH = "Kết quả KS Sale.xlsx"

if not os.path.exists(FILE_PATH):
    st.error(f"⚠️ Không tìm thấy file gốc tại `{FILE_PATH}`")
    st.stop()

try:
    with st.spinner("Đang kết xuất & xử lý thống kê..."):
         df = load_and_preprocess_data(FILE_PATH)
except Exception as e:
    st.error(f"Lỗi hệ thống khi đọc File: {e}")
    st.stop()

if df.empty:
    st.warning("⚠️ Bảng dữ liệu rỗng. Cần check file gốc.")
    st.stop()


# 2. HEADER
st.title("☁️ Dashboard Năng Suất & Kết Quả Khảo Sát Sale")
st.markdown("---")

# 3. SIDEBAR: LỌC LIÊN HOÀN (DYNAMIC CASCADING CROSS-FILTERS)
st.sidebar.markdown("### 🎯 Phễu Quản Trị")

def get_filter_opts(col_name, current_val, exclude_col=None):
    temp_df = df.copy()
    if exclude_col != 'Phòng' and st.session_state.get('sel_phong', 'Tất cả') != 'Tất cả':
        temp_df = temp_df[temp_df['Phòng'].astype(str) == st.session_state.get('sel_phong', 'Tất cả')]
    if exclude_col != 'Tháng' and st.session_state.get('sel_thang', 'Tất cả') != 'Tất cả':
        temp_df = temp_df[temp_df['Tháng'].astype(str) == st.session_state.get('sel_thang', 'Tất cả')]
    if exclude_col != 'Team' and st.session_state.get('sel_team', 'Tất cả') != 'Tất cả':
        temp_df = temp_df[temp_df['Team'].astype(str) == st.session_state.get('sel_team', 'Tất cả')]
    if exclude_col != 'Leader' and st.session_state.get('sel_leader', 'Tất cả') != 'Tất cả':
        temp_df = temp_df[temp_df['Leader'].astype(str) == st.session_state.get('sel_leader', 'Tất cả')]
    if exclude_col != 'Search_Name' and st.session_state.get('sel_nv', 'Tất cả') != 'Tất cả':
        temp_df = temp_df[temp_df['Search_Name'].astype(str) == st.session_state.get('sel_nv', 'Tất cả')]
        
    opts = ["Tất cả"] + sorted(list(temp_df[col_name].astype(str).unique()))
    
    idx = 0
    if current_val in opts:
        idx = opts.index(current_val)
    # Index=0 luôn là dự phòng 'Tất cả' để fix các lỗi xung đột State
    return opts, idx

# Render Widgets
phong_opts, phong_idx = get_filter_opts('Phòng', st.session_state.get('sel_phong', 'Tất cả'), 'Phòng')
selected_phong = st.sidebar.selectbox("🏢 1. Chọn Phòng", phong_opts, index=phong_idx, key='sel_phong')

thang_opts, thang_idx = get_filter_opts('Tháng', st.session_state.get('sel_thang', 'Tất cả'), 'Tháng')
selected_thang = st.sidebar.selectbox("📅 2. Chọn Tháng", thang_opts, index=thang_idx, key='sel_thang')

team_opts, team_idx = get_filter_opts('Team', st.session_state.get('sel_team', 'Tất cả'), 'Team')
selected_team = st.sidebar.selectbox("👥 3. Chọn Team", team_opts, index=team_idx, key='sel_team')

leader_opts, leader_idx = get_filter_opts('Leader', st.session_state.get('sel_leader', 'Tất cả'), 'Leader')
selected_leader = st.sidebar.selectbox("👨‍💼 4. Chọn Trưởng Nhóm", leader_opts, index=leader_idx, key='sel_leader')

nv_opts, nv_idx = get_filter_opts('Search_Name', st.session_state.get('sel_nv', 'Tất cả'), 'Search_Name')
selected_nv = st.sidebar.selectbox("👤 5. Chọn Nhân viên", nv_opts, index=nv_idx, key='sel_nv')

# Update Active DF
filtered_df = df.copy()
if selected_phong != 'Tất cả': filtered_df = filtered_df[filtered_df['Phòng'].astype(str) == selected_phong]
if selected_thang != 'Tất cả': filtered_df = filtered_df[filtered_df['Tháng'].astype(str) == selected_thang]
if selected_team != 'Tất cả': filtered_df = filtered_df[filtered_df['Team'].astype(str) == selected_team]
if selected_leader != 'Tất cả': filtered_df = filtered_df[filtered_df['Leader'].astype(str) == selected_leader]
if selected_nv != 'Tất cả': filtered_df = filtered_df[filtered_df['Search_Name'].astype(str) == selected_nv]


total_staff = len(filtered_df)

# 4. KHU VỰC KPI METRICS
st.subheader("📌 Tổng Quan KPI")

def calculate_percent(grade):
    if total_staff == 0: return 0
    return (filtered_df['Đánh giá_Clean'] == grade).sum() / total_staff * 100

col1, col2, col3, col4, col5 = st.columns(5)
col1.metric("Tổng Nhân Sự (Lọc)", f"{total_staff}")
col2.metric("🏆 Đạt Giỏi", f"{calculate_percent('Giỏi'):.1f}%")
col3.metric("🎯 Đạt Khá", f"{calculate_percent('Khá'):.1f}%")
col4.metric("⚖️ Trung Bình", f"{calculate_percent('Trung bình'):.1f}%")
col5.metric("🚨 Yếu", f"{calculate_percent('Yếu'):.1f}%")


# 5. CÂY LOGIC INSIGHTS NHẬN ĐỊNH TỰ ĐỘNG
st.markdown("### 💡 Nhận định tự động (Insights)")
def generate_insights(df_filtered, sel_team, sel_leader, sel_employee):
    staff_count = len(df_filtered)
    if staff_count == 0:
        st.info("Không đủ dữ liệu để tạo insights.")
        return
    
    # Tính cơ sở tỷ lệ toàn cục view hiện tại
    gioi_rate = (df_filtered['Đánh giá_Clean'] == 'Giỏi').sum() / staff_count * 100
    bad_rate = ((df_filtered['Đánh giá_Clean'] == 'Trung bình').sum() + (df_filtered['Đánh giá_Clean'] == 'Yếu').sum()) / staff_count * 100
    
    # TẦNG 1: CÁ NHÂN
    if sel_employee != "Tất cả":
        row = df_filtered.iloc[0]
        diem = row.get('Total points', 0)
        danh_gia = str(row.get('Đánh giá_Clean', 'Trống'))
        ten_nv = row.get('Họ và tên đầy đủ', sel_employee)
        
        if danh_gia in ['Giỏi', 'Khá']:
            st.success(f"🌟 **Biểu dương:** Nhân viên **{ten_nv}** đạt **{diem}** điểm (Xếp loại: **{danh_gia}**). Kỹ năng thao tác hệ thống RSA Ecom rất tốt!")
        elif danh_gia in ['Trung bình', 'Yếu']:
            st.warning(f"⚠️ **Lưu ý:** Nhân viên **{ten_nv}** đang xếp loại **{danh_gia}**. Cần training lại ngay các thao tác xử lý lỗi trên web/app và chính sách Voucher.")
        else:
            st.info(f"Nhân viên **{ten_nv}** xếp loại: {danh_gia}")
        return

    # TẦNG 2: LEADER
    if sel_leader != "Tất cả":
        khen = False
        if gioi_rate >= 50.0:
            st.success(f"🌟 **Biểu dương:** Nhóm của Leader **{sel_leader}** có tới **{gioi_rate:.1f}%** nhân sự đạt loại Giỏi.")
            khen = True
            
        if bad_rate >= 10.0:
            st.warning(f"⚠️ **Cảnh báo vận hành:** Nhóm của Leader **{sel_leader}** đang có tỷ lệ Trung bình/Yếu là **{bad_rate:.1f}%**. Leader cần rà soát lại kết quả đào tạo.")
            khen = True
            
        if not khen:
            st.info(f"Hiệu suất nhóm của Leader {sel_leader} ở mức ổn định chung.")
        return
            
    # TẦNG 3: TEAM
    if sel_team != "Tất cả":
        khen = False
        if gioi_rate >= 50.0:
            st.success(f"🌟 **Biểu dương:** Team **{sel_team}** có tới **{gioi_rate:.1f}%** nhân sự đạt loại Giỏi.")
            khen = True
            
        if bad_rate >= 10.0:
            st.warning(f"⚠️ **Cảnh báo vận hành:** Team **{sel_team}** đang có tỷ lệ Trung bình/Yếu là **{bad_rate:.1f}%**. Leader cần rà soát lại kết quả đào tạo.")
            khen = True
            
        if not khen:
            st.info(f"Hoạt động của Team {sel_team} ở mức ổn định, chưa có dấu hiệu đột phá hay rủi ro lớn.")
        return
        
    # TẦNG 4: TỔNG QUAN
    cat_types = ['Giỏi', 'Khá', 'Trung bình', 'Yếu']
    pivot_count = pd.crosstab(df_filtered['Team'], df_filtered['Đánh giá_Clean'])
    for c in cat_types:
        if c not in pivot_count.columns: pivot_count[c] = 0
    pivot_percent = pivot_count.div(pivot_count.sum(axis=1), axis=0) * 100
    
    if len(pivot_percent) > 0:
        if 'Giỏi' in pivot_percent.columns:
            top_gioi_teams = pivot_percent[pivot_percent['Giỏi'] >= 50.0]
            if not top_gioi_teams.empty:
                best_team = pivot_percent['Giỏi'].idxmax()
                best_val = pivot_percent['Giỏi'].max()
                st.success(f"🌟 **Biểu dương:** Team **{best_team}** đang dẫn đầu với **{best_val:.1f}%** nhân sự đạt loại Giỏi. Team đang nắm rất vững thao tác hệ thống và chính sách bán hàng!")
                
        pivot_percent['Bad_Rate'] = pivot_percent['Trung bình'] + pivot_percent['Yếu']
        bad_teams = pivot_percent[pivot_percent['Bad_Rate'] >= 10.0]
        if not bad_teams.empty:
            for team, row in bad_teams.iterrows():
                st.warning(f"⚠️ **Cảnh báo vận hành:** Team **{team}** đang có tỷ lệ Trung bình/Yếu lên tới **{row['Bad_Rate']:.1f}%**. Leader cần khẩn trương training lại cho team về các thao tác xử lý lỗi trên RSA Ecom, quy định áp dụng Voucher/Flash Sale và chính sách điểm thưởng để tránh sai sót đơn hàng.")
        
        if ('Giỏi' not in pivot_percent.columns or pivot_percent['Giỏi'].max() < 50.0) and bad_teams.empty:
            st.info("Hệ thống chưa tìm thấy dấu hiệu đột biến lớn nào từ bộ lọc hiện hành. Các Team đang giữ vững nhịp điệu.")

generate_insights(filtered_df, selected_team, selected_leader, selected_nv)

st.markdown("<hr>", unsafe_allow_html=True)


# 6. KHU VỰC PHÂN TÍCH TEAM LEVEL (CROSSTAB)
st.subheader("👥 Phân Tích Chất Lượng Nội Bộ Theo Team")
cat_types = ['Giỏi', 'Khá', 'Trung bình', 'Yếu']
pivot_count = pd.crosstab(filtered_df['Team'], filtered_df['Đánh giá_Clean'])
for c in cat_types:
    if c not in pivot_count.columns: pivot_count[c] = 0
pivot_count = pivot_count[['Giỏi', 'Khá', 'Trung bình', 'Yếu']]
pivot_percent = pivot_count.div(pivot_count.sum(axis=1), axis=0) * 100

if total_staff > 0:
    col_table1, col_table2 = st.columns(2)
    
    with col_table1:
        st.markdown("**1. Bảng Số lượng nhân sự (Headcount)**")
        df_count = pivot_count.copy()
        df_count['(Tổng HS)'] = df_count.sum(axis=1)
        st.table(df_count)
        
    with col_table2:
        st.markdown("**2. Bảng Tỷ lệ phần trăm trong Team (%)**")
        df_percent = pivot_percent[['Giỏi', 'Khá', 'Trung bình', 'Yếu']].copy().round(1)
        for c in cat_types:
            df_percent[c] = df_percent[c].astype(str) + " %"
        st.table(df_percent)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # 7. KHU VỰC BIỂU ĐỒ
    st.subheader("📊 Biểu Đồ Cơ Cấu Chất Lượng Đánh Giá Theo Team")
    
    melted_df = pivot_count.reset_index().melt(id_vars='Team', var_name='Đánh giá', value_name='Số lượng')
    
    color_map = {
        'Giỏi': '#2ca02c',       
        'Khá': '#1f77b4',        
        'Trung bình': '#ff7f0e', 
        'Yếu': '#d62728'         
    }
    
    fig = px.bar(melted_df, x='Team', y='Số lượng', color='Đánh giá', 
                    color_discrete_map=color_map, 
                    hover_data={'Số lượng': True})
    
    fig.update_layout(
        barmode='stack', 
        barnorm='percent', 
        xaxis_title=None, 
        yaxis_title="Tỷ lệ %",
        paper_bgcolor='rgba(0,0,0,0)',
        plot_bgcolor='rgba(0,0,0,0)',
        margin=dict(t=20, l=10, r=10, b=10)
    )
    fig.update_xaxes(showgrid=False, tickfont=dict(size=14, weight='bold', color='black'))
    fig.update_yaxes(showgrid=False, ticks="outside", ticksuffix="%")
    
    st.plotly_chart(fig, use_container_width=True)

else:
    st.info("Trống dữ liệu phân tích.")

st.markdown("<hr>", unsafe_allow_html=True)


# 8. DATA TABLE
st.subheader("📋 Danh Sách Chi Tiết Nhân Viên")

if total_staff > 0:
    base_cols = ['STT', 'Tháng', 'Phòng', 'Search_Name', 'Mail FPT', 'Leader', 'Team', 'Total points', 'Đánh giá']
    show_cols = [c for c in base_cols if c in filtered_df.columns]
    
    st.dataframe(filtered_df[show_cols], use_container_width=True, hide_index=True)
else:
    st.info("Không có dữ liệu chi tiết phù hợp với bộ lọc hiện tại.")
