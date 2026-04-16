import streamlit as st
import pandas as pd
import plotly.express as px
import os

# 1. CẤU HÌNH TRANG VÀ CSS (1 TRANG DUY NHẤT)
st.set_page_config(page_title="Dashboard Năng Suất & Kết Quả Khảo Sát Sale", layout="wide", initial_sidebar_state="expanded")

st.markdown("""
    <style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    
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
    
    /* Dark mode adapt */
    @media (prefers-color-scheme: dark) {
        div[data-testid="metric-container"] {
            background-color: #1e1e1e;
            border: 1px solid #333333;
            box-shadow: 0px 4px 10px rgba(0, 0, 0, 0.6);
        }
        div[data-testid="metric-container"] label {
            color: #dddddd;
        }
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
    # Đọc Sheet đầu tiên, giới hạn cột A->J
    df = pd.read_excel(file_path, sheet_name=0, usecols="A:J")
    df.columns = df.columns.astype(str).str.strip()
    
    # Ép kiểu 'Inside' thành chuỗi String (Loại bỏ ArrowInvalid)
    if 'Inside' in df.columns:
        df['Inside'] = df['Inside'].astype(str).str.replace(r'\.0$', '', regex=True).str.strip()
        df['Inside'] = df['Inside'].replace('nan', 'Unknown')
    else:
        df['Inside'] = 'Unknown'
        
    if 'Mail FPT' in df.columns:
        df['Mail FPT'] = df['Mail FPT'].astype(str).replace('nan', '')
        
    if 'Họ và tên đầy đủ' in df.columns:
        df['Họ và tên đầy đủ'] = df['Họ và tên đầy đủ'].astype(str).str.strip()
    
    # Tạo Search_Name
    df['Search_Name'] = df['Inside'] + " - " + df.get('Họ và tên đầy đủ', '')
    
    # Chuẩn hóa dữ liệu missing
    df['Phòng'] = df.get('Phòng', pd.Series(['Unknown'] * len(df))).fillna('Unknown')
    df['Tháng'] = df.get('Tháng', pd.Series(['Unknown'] * len(df))).astype(str).fillna('Unknown')
    df['Team'] = df.get('Team', pd.Series(['Unknown'] * len(df))).fillna('Unknown')
    df['Leader'] = df.get('Leader', pd.Series(['Unknown'] * len(df))).fillna('Unknown')
    df['Đánh giá'] = df.get('Đánh giá', pd.Series(['Trống'] * len(df))).fillna('Trống')
    
    # Mặc định chữ clean
    df['Đánh giá_Clean'] = df['Đánh giá'].str.capitalize()
    
    # Clean data (Loại rác tính tổng)
    if 'Họ và tên đầy đủ' in df.columns:
        df = df[~df['Họ và tên đầy đủ'].str.contains('Tổng|Total|Average', case=False, na=False)]
        df = df[df['Họ và tên đầy đủ'] != 'nan']

    return df


# === ĐỌC DỮ LIỆU ===
FILE_PATH = r"c:\Users\User\Documents\ks\Kết quả KS Sale.xlsx"

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

# 3. SIDEBAR: CASCADE
st.sidebar.markdown("### 🎯 Phễu Quản Trị")

phong_list = ["Tất cả"] + sorted(list(df['Phòng'].astype(str).unique()))
selected_phong = st.sidebar.selectbox("🏢 1. Chọn Phòng", phong_list)

f_df = df.copy()
if selected_phong != "Tất cả":
    f_df = f_df[f_df['Phòng'].astype(str) == selected_phong]

thang_list = ["Tất cả"] + sorted(list(f_df['Tháng'].astype(str).unique()))
selected_thang = st.sidebar.selectbox("📅 2. Chọn Tháng", thang_list)

if selected_thang != "Tất cả":
    f_df = f_df[f_df['Tháng'].astype(str) == selected_thang]

team_list = ["Tất cả"] + sorted(list(f_df['Team'].astype(str).unique()))
selected_team = st.sidebar.selectbox("👥 3. Chọn Team", team_list)

if selected_team != "Tất cả":
    f_df = f_df[f_df['Team'].astype(str) == selected_team]

leader_list = ["Tất cả"] + sorted(list(f_df['Leader'].astype(str).unique()))
selected_leader = st.sidebar.selectbox("👨‍💼 4. Chọn Trưởng Nhóm", leader_list)

if selected_leader != "Tất cả":
    f_df = f_df[f_df['Leader'].astype(str) == selected_leader]
    
nv_list = ["Tất cả"] + sorted(list(f_df['Search_Name'].astype(str).unique()))
selected_nv = st.sidebar.selectbox("👤 5. Chọn Nhân viên", nv_list)

if selected_nv != "Tất cả":
    f_df = f_df[f_df['Search_Name'].astype(str) == selected_nv]

filtered_df = f_df.copy()

# TÍNH TOÁN NGẦM ĐỂ TẠO INSIGHTS LUÔN
total_staff = len(filtered_df)
cat_types = ['Giỏi', 'Khá', 'Trung bình', 'Yếu']
pivot_count = pd.crosstab(filtered_df['Team'], filtered_df['Đánh giá_Clean'])
for c in cat_types:
    if c not in pivot_count.columns:
        pivot_count[c] = 0
pivot_count = pivot_count[['Giỏi', 'Khá', 'Trung bình', 'Yếu']]
pivot_percent = pivot_count.div(pivot_count.sum(axis=1), axis=0) * 100


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

# 5. AUTOMATED INSIGHTS
st.markdown("### 💡 Nhận định tự động (Insights)")
if total_staff > 0 and len(pivot_percent) > 0:
    # Lấy top team Giỏi
    if 'Giỏi' in pivot_percent.columns and pivot_percent['Giỏi'].max() > 0:
        top_team_gioi = pivot_percent['Giỏi'].idxmax()
        max_gioi_val = pivot_percent['Giỏi'].max()
        st.success(f"🌟 **Biểu dương:** Team **{top_team_gioi}** đang dẫn đầu với **{max_gioi_val:.1f}%** nhân sự đạt loại Giỏi. Team đang nắm rất vững thao tác hệ thống và chính sách bán hàng!")
    
    # Lấy các team bị cảnh báo (TB + Yếu >= 10%)
    pivot_percent['Bad_Rate'] = pivot_percent['Trung bình'] + pivot_percent['Yếu']
    bad_teams = pivot_percent[pivot_percent['Bad_Rate'] >= 10.0]
    
    if not bad_teams.empty:
        for team, row in bad_teams.iterrows():
            st.warning(f"⚠️ **Cảnh báo vận hành:** Team **{team}** đang có tỷ lệ Trung bình/Yếu lên tới **{row['Bad_Rate']:.1f}%**. Leader cần khẩn trương training lại cho team về các thao tác xử lý lỗi trên RSA Ecom, quy định áp dụng Voucher/Flash Sale và chính sách điểm thưởng để tránh sai sót đơn hàng.")
            
    if (not 'Giỏi' in pivot_percent.columns or pivot_percent['Giỏi'].max() == 0) and bad_teams.empty:
        st.info("Hệ thống chưa tìm thấy dấu hiệu đột biến lớn nào từ bộ lọc hiện hành.")
else:
    st.info("Không đủ dữ liệu để tạo insights.")

st.markdown("<hr>", unsafe_allow_html=True)


# 6. KHU VỰC PHÂN TÍCH TEAM LEVEL
st.subheader("👥 Phân Tích Chất Lượng Nội Bộ Theo Team")

if total_staff > 0:
    col_table1, col_table2 = st.columns(2)
    
    with col_table1:
        st.markdown("**1. Bảng Số lượng nhân sự (Headcount)**")
        df_count = pivot_count.copy()
        df_count['(Tổng HS)'] = df_count.sum(axis=1)
        st.table(df_count)
        
    with col_table2:
        st.markdown("**2. Bảng Tỷ lệ phần trăm trong Team (%)**")
        df_percent = pivot_percent[['Giỏi', 'Khá', 'Trung bình', 'Yếu']].copy().round(1) # Drop Bad_Rate out
        for c in cat_types:
            df_percent[c] = df_percent[c].astype(str) + " %"
        st.table(df_percent)

    st.markdown("<br>", unsafe_allow_html=True)
    
    # 7. KHU VỰC BIỂU ĐỒ NÂNG CẤP
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
    # Tăng độ đậm trục X
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
