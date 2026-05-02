import streamlit as st
import pandas as pd
import plotly.express as px
import os

st.set_page_config(page_title="Dashboard Analisis", layout="wide")

# ==========================================
# 1. Custom CSS (Minimalist & Professional)
# ==========================================
st.markdown("""
<style>
    html, body, [class*="css"] { font-family: 'Inter', 'Helvetica Neue', Helvetica, Arial, sans-serif; }
    .stApp { background-color: #FAFAFA; }
    [data-testid="stSidebar"] { background-color: #FFFFFF; border-right: 1px solid #EAEAEA; }
    h1 { color: #111827; font-weight: 600; font-size: 2rem; padding-bottom: 0.5rem; }
    .subtitle { color: #6B7280; font-size: 1rem; margin-bottom: 2rem; }
    div[data-testid="metric-container"] { background-color: #FFFFFF; border: 1px solid #E5E7EB; padding: 1.25rem; border-radius: 4px; }
    div[data-testid="metric-container"] label { color: #6B7280 !important; font-weight: 500 !important; font-size: 0.875rem !important; }
    div[data-testid="metric-container"] div[data-testid="stMetricValue"] { color: #111827 !important; font-size: 1.5rem !important; font-weight: 600 !important; }
    h3 { color: #374151; font-size: 1.1rem; font-weight: 600; margin-top: 1rem; margin-bottom: 0.5rem; }
</style>
""", unsafe_allow_html=True)

# ==========================================
# 2. Data Loading
# ==========================================
@st.cache_data
def load_data():
    file_path = 'Southwind (1).xlsx'
    if not os.path.exists(file_path):
        st.error(f"File '{file_path}' tidak ditemukan.")
        st.stop()
        
    try:
        orders = pd.read_excel(file_path, sheet_name='Orders')
        people = pd.read_excel(file_path, sheet_name='People')
        returns = pd.read_excel(file_path, sheet_name='Returns')
        
        people = people.rename(columns={'Region': 'wilayah'})
        returns = returns.rename(columns={'Order ID': 'id_pemesanan'})
        
        df = orders.merge(people, on='wilayah', how='left')
        df = df.merge(returns, on='id_pemesanan', how='left')
        df['Returned'] = df['Returned'].fillna('No')
        
        if 'tanggal_pemesanan' in df.columns:
            df['tanggal_pemesanan'] = pd.to_datetime(df['tanggal_pemesanan'])
            df['Bulan_Tahun'] = df['tanggal_pemesanan'].dt.to_period('M').astype(str)
            
        if 'kodepos' in df.columns:
            df['kodepos'] = df['kodepos'].astype(str)
            
        return df
    except Exception as e:
        st.error(f"Terjadi kesalahan: {e}")
        st.stop()

# ==========================================
# 3. Sidebar & Slicers
# ==========================================
df = load_data()
with st.sidebar:
    st.markdown("<h3 style='margin-top: 0;'>⚙️ Slicers & Filter</h3>", unsafe_allow_html=True)
    
    # 1. Date Range Slicer
    if 'tanggal_pemesanan' in df.columns:
        min_date = df['tanggal_pemesanan'].min().date()
        max_date = df['tanggal_pemesanan'].max().date()
        
        date_selection = st.date_input(
            "Rentang Waktu",
            value=(min_date, max_date),
            min_value=min_date,
            max_value=max_date
        ) 
    st.markdown("---")
    
    # 2. Category & Region Filters
    wilayah_list = ['Semua'] + list(df['wilayah'].dropna().unique())
    kategori_list = ['Semua'] + list(df['Kategori_produk'].dropna().unique())
    segmen_list = ['Semua'] + list(df['segmen'].dropna().unique())
    
    wilayah_filter = st.selectbox("Wilayah", options=wilayah_list)
    kategori_filter = st.selectbox("Kategori Produk", options=kategori_list)
    segmen_filter = st.selectbox("Segmen Pelanggan", options=segmen_list)

# Apply filters
df_selection = df.copy()
if 'tanggal_pemesanan' in df_selection.columns and len(date_selection) == 2:
    start_date, end_date = date_selection
    df_selection = df_selection[
        (df_selection['tanggal_pemesanan'].dt.date >= start_date) & 
        (df_selection['tanggal_pemesanan'].dt.date <= end_date)
    ]
if wilayah_filter != 'Semua':
    df_selection = df_selection[df_selection['wilayah'] == wilayah_filter]
if kategori_filter != 'Semua':
    df_selection = df_selection[df_selection['Kategori_produk'] == kategori_filter]
if segmen_filter != 'Semua':
    df_selection = df_selection[df_selection['segmen'] == segmen_filter]
if df_selection.empty:
    st.warning("Data kosong untuk filter yang dipilih.")
    st.stop()
# Tambahkan kolom untuk indikator profit (Positif vs Negatif)
df_selection['Status Profit'] = df_selection['keuntungan'].apply(lambda x: 'Untung' if x > 0 else 'Rugi')

# ==========================================
# 4. Main Dashboard Layout
# ==========================================
st.markdown("<h1>Dashboard Analisis Penjualan & Keuntungan</h1>", unsafe_allow_html=True)

total_sales = df_selection['Penjualan'].sum()
total_profit = df_selection['keuntungan'].sum()
total_orders = len(df_selection)
return_count = len(df_selection[df_selection['Returned'] == 'Yes'])
return_rate = (return_count / total_orders * 100) if total_orders > 0 else 0

col_kpi1, col_kpi2, col_kpi3, col_kpi4 = st.columns(4)
with col_kpi1: st.metric("Total Penjualan", f"${total_sales:,.0f}")
with col_kpi2: st.metric("Total Keuntungan", f"${total_profit:,.0f}")
with col_kpi3: st.metric("Total Pesanan", f"{total_orders:,}")
with col_kpi4: st.metric("Tingkat Retur", f"{return_rate:.1f}%")

st.markdown("<br>", unsafe_allow_html=True)

# Custom Colors
color_profit = '#10B981' # Hijau (Untung)
color_loss = '#EF4444'   # Merah (Rugi)
color_sales = '#3B82F6'  # Biru

# ==========================================
# VISUALISASI
# ==========================================

col1, col2 = st.columns(2)

# 1. Visualisasi Penjualan per Kategori 
with col1:
    st.markdown("### 1. Penjualan per Kategori Produk")
    cat_df = df_selection.groupby('Kategori_produk')['Penjualan'].sum().reset_index().sort_values('Penjualan')
    fig1 = px.bar(cat_df, x='Penjualan', y='Kategori_produk', color='Kategori_produk', orientation='h',
                  color_discrete_sequence=px.colors.qualitative.Bold)
    fig1.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", showlegend=False,
        xaxis=dict(showgrid=True, gridcolor='#E5E7EB', tickprefix='$'),
        yaxis=dict(showgrid=False), margin=dict(t=20, b=20, l=10, r=10)
    )
    st.plotly_chart(fig1, use_container_width=True)

# 2. Visualisasi Penjualan per Wilayah 
with col2:
    st.markdown("### 2. Penjualan per Wilayah")
    wil_df = df_selection.groupby('wilayah')['Penjualan'].sum().reset_index().sort_values('Penjualan', ascending=False)
    fig2 = px.bar(wil_df, x='wilayah', y='Penjualan', color='wilayah',
                  color_discrete_sequence=px.colors.qualitative.Safe)
    fig2.update_layout(
        plot_bgcolor="rgba(0,0,0,0)", showlegend=False,
        yaxis=dict(showgrid=True, gridcolor='#E5E7EB', tickprefix='$'),
        xaxis=dict(showgrid=False), margin=dict(t=20, b=20, l=10, r=10)
    )
    st.plotly_chart(fig2, use_container_width=True)

st.markdown("---")

# 3. Visualisasi Tren Penjualan per kategori produk
st.markdown("### 3. Tren Penjualan per Kategori Produk (Waktu ke Waktu)")
if 'Bulan_Tahun' in df_selection.columns:
    trend_df = df_selection.groupby(['Bulan_Tahun', 'Kategori_produk'])['Penjualan'].sum().reset_index()
    fig3 = px.line(trend_df, x='Bulan_Tahun', y='Penjualan', color='Kategori_produk', 
                   color_discrete_sequence=px.colors.qualitative.Bold)
    fig3.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        yaxis=dict(showgrid=True, gridcolor='#E5E7EB', tickprefix='$'),
        xaxis=dict(showgrid=False), margin=dict(t=20, b=20, l=10, r=10), legend_title_text=''
    )
    st.plotly_chart(fig3, use_container_width=True)

st.markdown("---")

col3, col4 = st.columns(2)

# 4. Visualisasi Segmentasi Pelanggan
with col3:
    st.markdown("### 4. Segmentasi Pelanggan (Kontribusi Penjualan)")
    seg_df = df_selection.groupby('segmen')['Penjualan'].sum().reset_index()
    fig4 = px.pie(seg_df, names='segmen', values='Penjualan', hole=0.4, 
                  color_discrete_sequence=px.colors.qualitative.Vivid)
    fig4.update_layout(margin=dict(t=20, b=20, l=10, r=10))
    fig4.update_traces(hovertemplate="%{label}<br>Penjualan: $%{value:,.0f}<br>%{percent}")
    st.plotly_chart(fig4, use_container_width=True)

# 5. Visualisasi Pengaruh Diskon terhadap Keuntungan
# Mewarnai berdasarkan Untung (Hijau) atau Rugi (Merah)
with col4:
    st.markdown("### 5. Pengaruh Diskon terhadap Keuntungan")
    fig5 = px.scatter(df_selection, x='diskon', y='keuntungan', color='Status Profit',
                      color_discrete_map={'Untung': color_profit, 'Rugi': color_loss}, opacity=0.6)
    fig5.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title='Diskon', showgrid=True, gridcolor='#E5E7EB'),
        yaxis=dict(title='Keuntungan', showgrid=True, gridcolor='#E5E7EB', tickprefix='$'),
        margin=dict(t=20, b=20, l=10, r=10), legend_title_text=''
    )
    st.plotly_chart(fig5, use_container_width=True)

st.markdown("---")

col5, col6 = st.columns(2)

# 6. Visualisasi Analisis Penjualan vs Keuntungan
# Mewarnai berdasarkan Untung/Rugi agar sangat jelas apakah penjualan tinggi = selalu untung
with col5:
    st.markdown("### 6. Analisis Penjualan vs Keuntungan")
    fig6 = px.scatter(df_selection, x='Penjualan', y='keuntungan', color='Status Profit',
                      color_discrete_map={'Untung': color_profit, 'Rugi': color_loss}, opacity=0.6)
    fig6.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title='Penjualan', showgrid=True, gridcolor='#E5E7EB', tickprefix='$'),
        yaxis=dict(title='Keuntungan', showgrid=True, gridcolor='#E5E7EB', tickprefix='$'),
        margin=dict(t=20, b=20, l=10, r=10), legend_title_text=''
    )
    st.plotly_chart(fig6, use_container_width=True)

# 7. Visualisasi Analisis Subkategori
# Biru untuk Penjualan, Hijau untuk Keuntungan
with col6:
    st.markdown("### 7. Analisis Subkategori (Penjualan vs Keuntungan)")
    subcat_df = df_selection.groupby('sub_kategori_produk')[['Penjualan', 'keuntungan']].sum().reset_index()
    subcat_melted = subcat_df.melt(id_vars='sub_kategori_produk', value_vars=['Penjualan', 'keuntungan'], 
                                   var_name='Metrik', value_name='Nilai')
    
    subcat_order = subcat_df.sort_values('Penjualan', ascending=False)['sub_kategori_produk'].tolist()
    
    fig7 = px.bar(subcat_melted, x='sub_kategori_produk', y='Nilai', color='Metrik', 
                  barmode='group', color_discrete_map={'Penjualan': color_sales, 'keuntungan': color_profit},
                  category_orders={"sub_kategori_produk": subcat_order})
    fig7.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis=dict(title='', showgrid=False),
        yaxis=dict(title='Nilai ($)', showgrid=True, gridcolor='#E5E7EB', tickprefix='$'),
        margin=dict(t=20, b=20, l=10, r=10), legend_title_text=''
    )
    st.plotly_chart(fig7, use_container_width=True)

st.markdown("<hr style='margin-top: 2rem; margin-bottom: 2rem;'>", unsafe_allow_html=True)

# ==========================================
# 8. INFORMASI / INSIGHT PENTING
# ==========================================
st.markdown("### 8. Ringkasan Insight & Temuan Penting")

if not df_selection.empty:
    # Insight 1: Wilayah dengan Penjualan Terbaik
    best_region = df_selection.groupby('wilayah')['Penjualan'].sum().idxmax()
    best_region_sales = df_selection.groupby('wilayah')['Penjualan'].sum().max()
    
    # Insight 2: Produk Paling Menguntungkan
    best_product = df_selection.groupby('nama_produk')['keuntungan'].sum().idxmax()
    best_product_profit = df_selection.groupby('nama_produk')['keuntungan'].sum().max()
    
    # Insight 3: Pengaruh Diskon terhadap Profit (Korelasi)
    if 'diskon' in df_selection.columns and df_selection['diskon'].nunique() > 1:
        corr_discount_profit = df_selection['diskon'].corr(df_selection['keuntungan'])
        if pd.isna(corr_discount_profit):
            discount_insight = "Data diskon seragam, tidak dapat menyimpulkan pengaruhnya terhadap keuntungan."
        elif corr_discount_profit < -0.1:
            discount_insight = "Terdapat **korelasi negatif** yang cukup jelas. Artinya, semakin tinggi diskon yang diberikan, keuntungan cenderung **menurun** atau memicu kerugian."
        elif corr_discount_profit > 0.1:
            discount_insight = "Terdapat **korelasi positif**. Pemberian diskon terbukti mampu mendongkrak keuntungan secara keseluruhan."
        else:
            discount_insight = "Tidak terlihat pengaruh linier yang signifikan antara besaran diskon dan keuntungan yang didapat."
    else:
        discount_insight = "Data diskon tidak cukup bervariasi untuk menyimpulkan korelasinya dengan keuntungan."

    # Insight 4: Segmen Pelanggan Paling Menguntungkan
    best_segment = df_selection.groupby('segmen')['keuntungan'].sum().idxmax()
    best_segment_profit = df_selection.groupby('segmen')['keuntungan'].sum().max()
    
    # Insight 5: Kategori Produk Paling Sering Diretur
    retur_df = df_selection[df_selection['Returned'] == 'Yes']
    if not retur_df.empty:
        worst_return_cat = retur_df.groupby('Kategori_produk').size().idxmax()
        worst_return_count = retur_df.groupby('Kategori_produk').size().max()
        return_insight = f"Kategori **{worst_return_cat}** mencatat pengembalian barang terbanyak ({worst_return_count} pesanan diretur)."
    else:
        return_insight = "Tidak ada riwayat retur barang pada filter saat ini."

    st.info(f'''
    Berdasarkan data yang difilter saat ini, berikut adalah 5 temuan utama:
    
    1. **Wilayah Penjualan Terbaik**: Wilayah **{best_region}** mendominasi dengan total penjualan tertinggi mencapai **${best_region_sales:,.0f}**.
    2. **Produk Paling Menguntungkan**: Produk **{best_product}** merupakan kontributor profit terbesar dengan total keuntungan **${best_product_profit:,.0f}**.
    3. **Segmen Pelanggan Terbaik**: Segmen **{best_segment}** memberikan total keuntungan tertinggi sebesar **${best_segment_profit:,.0f}**.
    4. **Analisis Barang Retur**: {return_insight}
    5. **Pengaruh Diskon terhadap Keuntungan**: {discount_insight}
    ''')

st.markdown("<hr style='margin-top: 2rem; margin-bottom: 2rem;'>", unsafe_allow_html=True)

# ---- Data Table ----
st.markdown("### Detail Data Transaksi")
kolom_penting = ['id_pemesanan', 'tanggal_pemesanan', 'nama_pelanggan', 'wilayah', 'nama_produk', 'Kategori_produk', 'Penjualan', 'keuntungan', 'Status Profit']
kolom_tersedia = [col for col in kolom_penting if col in df_selection.columns]

st.dataframe(
    df_selection[kolom_tersedia], 
    use_container_width=True,
    height=300,
    hide_index=True
)