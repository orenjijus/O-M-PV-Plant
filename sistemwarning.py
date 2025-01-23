import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px  # Untuk visualisasi interaktif
import os

# ---------------------------
# Konstanta Utama
# ---------------------------
PV_CAPACITY = 2.06               # MWp (kapasitas PV plant)
PR_THRESHOLD = 0.75              # Batas Performance Ratio
INVERTER_EFF_THRESHOLD = 0.90    # Batas efisiensi Inverter

# ==================================================
# Fungsi-Fungsi Bantuan (Sama seperti contoh sebelumnya)
# ==================================================
def load_sensor_data_em(file_em):
    """
    Membaca file Excel EM, mengatur kolom secara dinamis,
    dan memproses data agar siap dipakai.
    """
    df_em = pd.ExcelFile(file_em).parse('5 minutes')
    em_columns = df_em.iloc[2].tolist()
    
    df_em.columns = em_columns
    df_em = df_em.iloc[3:].rename(columns={
        em_columns[3]: 'Start Time',
        em_columns[4]: 'Irradiance'
    })
    
    df_em['Irradiance'] = pd.to_numeric(df_em['Irradiance'], errors='coerce')
    df_em = df_em.dropna(subset=['Irradiance'])
    df_em = df_em[df_em['Irradiance'] >= 0]
    
    return df_em

def load_revenue_meter_data_rm(file_rm):
    """
    Membaca file Excel RM, mengatur kolom secara dinamis,
    dan memproses data agar siap dipakai.
    """
    df_rm = pd.ExcelFile(file_rm).parse('5 minutes')
    rm_columns = df_rm.iloc[2].tolist()
    
    df_rm.columns = rm_columns
    df_rm = df_rm.iloc[3:].rename(columns={
        rm_columns[3]: 'Start Time',
        rm_columns[5]: 'Active Energy (kWh)'
    })
    
    df_rm['Active Energy (kWh)'] = pd.to_numeric(df_rm['Active Energy (kWh)'], errors='coerce')
    df_rm = df_rm.dropna(subset=['Active Energy (kWh)'])
    df_rm = df_rm[df_rm['Active Energy (kWh)'] >= 0]
    
    return df_rm

def calculate_performance_ratio(df_em, df_rm, pv_capacity=PV_CAPACITY, threshold=PR_THRESHOLD):
    """
    Menggabungkan data EM & RM, hitung PR, dan memberi label performa.
    """
    merged_df = pd.merge(df_em, df_rm, on='Start Time', how='inner')
    merged_df['PR'] = merged_df['Active Energy (kWh)'] / (merged_df['Irradiance'] * pv_capacity * 1000)
    merged_df['Performance Status'] = merged_df['PR'].apply(
        lambda x: 'Good' if x >= threshold else 'Needs Attention'
    )
    return merged_df

def identify_performance_issues(df, threshold=PR_THRESHOLD):
    """
    Identifikasi masalah performa jika PR < threshold.
    """
    def check_issue(row):
        if row['PR'] < threshold:
            if row['PR'] < threshold * 0.9:
                return 'Kotoran Modul'
            else:
                return 'Kalibrasi Sensor Dibutuhkan'
        return 'Tidak Ada Masalah'
    
    df['Indikasi Masalah'] = df.apply(check_issue, axis=1)
    return df

def analyze_inverter_performance(merged_df, inverter_df_list, pv_capacity=PV_CAPACITY,
                                 eff_threshold=INVERTER_EFF_THRESHOLD):
    """
    Menganalisis performa beberapa inverter dengan data frame
    (sudah di-load di memori, bukan file).
    """
    results = []
    for inv_data, inv_file_name in inverter_df_list:
        # Pastikan kolom sesuai
        df_merged_inv = pd.merge(merged_df, inv_data, on='Start Time', how='inner')
        
        # Simulated energy
        df_merged_inv['Simulated Energy (kWh)'] = df_merged_inv['Irradiance'] * pv_capacity * 1000
        
        df_merged_inv = df_merged_inv[df_merged_inv['Simulated Energy (kWh)'] > 0]
        df_merged_inv['Inverter Efficiency'] = df_merged_inv['Energy Output (kWh)'] / df_merged_inv['Simulated Energy (kWh)']
        
        low_eff_df = df_merged_inv[df_merged_inv['Inverter Efficiency'] < eff_threshold]
        results.append({
            'Inverter File': inv_file_name,
            'Low Efficiency Count': len(low_eff_df),
            'Mean Efficiency': df_merged_inv['Inverter Efficiency'].mean()
        })
    
    summary_df = pd.DataFrame(results)
    return summary_df

def load_inverter_data(file):
    """
    Load data inverter dari file Excel,
    lalu kembalikan DataFrame dengan kolom
    'Start Time' dan 'Energy Output (kWh)'.
    """
    inv_data = pd.ExcelFile(file).parse('5 minutes')
    inv_columns = inv_data.iloc[2].tolist()
    inv_data.columns = inv_columns
    inv_data = inv_data.iloc[3:].rename(columns={
        inv_columns[3]: 'Start Time',
        inv_columns[5]: 'Energy Output (kWh)'
    })
    
    inv_data['Energy Output (kWh)'] = pd.to_numeric(inv_data['Energy Output (kWh)'], errors='coerce')
    inv_data = inv_data.dropna(subset=['Energy Output (kWh)'])
    inv_data = inv_data[inv_data['Energy Output (kWh)'] >= 0]
    return inv_data

# ================================
# Bagian Aplikasi Streamlit
# ================================

def main():
    st.title("Sistem Analisis Performa PV Plant")
    st.markdown("""
    **Tahap 1**: Analisis PR (Makro Level)  
    **Tahap 2**: Identifikasi Masalah PR  
    **Tahap 3**: Analisis Performa Inverter  
    ---
    """)

    # ----------------------------------------
    # 1. Upload File EM & RM
    # ----------------------------------------
    st.sidebar.header("Upload Data EM & RM")
    em_file = st.sidebar.file_uploader("Upload File EM (Excel)", type=['xlsx'])
    rm_file = st.sidebar.file_uploader("Upload File RM (Excel)", type=['xlsx'])

    # List inverter file uploads
    st.sidebar.header("Upload File Inverter")
    inverter_files = st.sidebar.file_uploader("Upload Multiple Inverter Files", type=['xlsx'], accept_multiple_files=True)

    # Tombol proses
    if st.sidebar.button("Proses Analisis"):
        # Pastikan file ada
        if (em_file is not None) and (rm_file is not None):
            # 2. Load Data
            with st.spinner("Mengupload & memproses data EM..."):
                df_em = load_sensor_data_em(em_file)
            with st.spinner("Mengupload & memproses data RM..."):
                df_rm = load_revenue_meter_data_rm(rm_file)
            
            st.success("Data EM & RM berhasil diproses.")

            # 3. Hitung PR - Tahap 1
            merged_data = calculate_performance_ratio(df_em, df_rm, PV_CAPACITY, PR_THRESHOLD)
            # 4. Identifikasi Masalah - Tahap 2
            merged_data = identify_performance_issues(merged_data, PR_THRESHOLD)

            # Tampilkan ringkasan
            st.subheader("Ringkasan Hasil (Tahap 1 & 2)")
            st.write("Berikut beberapa baris data hasil:")
            st.dataframe(merged_data.head(10))
            
            # Statistik PR
            st.write("**Statistik PR**:")
            st.write(merged_data['PR'].describe())
            
            # Visualisasi PR (misalnya plot seiring waktu)
            st.write("**Grafik PR vs. Waktu**")
            # Pastikan 'Start Time' terbaca sebagai datetime jika ingin plot
            # (Tergantung format data aslinya)
            # Jika 'Start Time' bukan format datetime, konversi:
            merged_data['Start Time'] = pd.to_datetime(merged_data['Start Time'], errors='coerce')
            
            fig_pr_time = px.line(
                merged_data.sort_values('Start Time'),
                x='Start Time',
                y='PR',
                title='Performance Ratio (PR) over Time'
            )
            st.plotly_chart(fig_pr_time, use_container_width=True)
            
            # Pie chart atau bar chart untuk status:
            st.write("**Distribusi Performance Status**")
            status_count = merged_data['Performance Status'].value_counts().reset_index()
            status_count.columns = ['Status', 'Count']
            fig_status = px.bar(
                status_count, x='Status', y='Count', 
                title='Jumlah Data Good vs Needs Attention',
                color='Status'
            )
            st.plotly_chart(fig_status, use_container_width=True)
            
            # Pie chart Indikasi Masalah
            st.write("**Indikasi Masalah** (untuk PR yang < 0.75)")
            issue_count = merged_data[merged_data['PR'] < PR_THRESHOLD]['Indikasi Masalah'].value_counts().reset_index()
            issue_count.columns = ['Issue', 'Count']
            fig_issue = px.pie(issue_count, values='Count', names='Issue', title='Indikasi Masalah PR Rendah')
            st.plotly_chart(fig_issue, use_container_width=True)

            # ----------------------------------------
            # 5. Analisis Inverter - Tahap 3
            # ----------------------------------------
            if inverter_files:
                inverter_df_list = []
                for inv_file in inverter_files:
                    inv_data = load_inverter_data(inv_file)
                    # Gunakan nama file nya untuk identifikasi
                    inverter_df_list.append((inv_data, inv_file.name))
                
                # Jalankan analisis
                with st.spinner("Menganalisis data inverter..."):
                    inverter_summary = analyze_inverter_performance(merged_data, inverter_df_list,
                                                                    pv_capacity=PV_CAPACITY,
                                                                    eff_threshold=INVERTER_EFF_THRESHOLD)
                st.success("Analisis Inverter Selesai.")
                
                st.subheader("Ringkasan Performa Inverter (Tahap 3)")
                st.dataframe(inverter_summary)

                # Visualisasi ringkasan inverter efficiency
                if 'Mean Efficiency' in inverter_summary.columns:
                    fig_eff = px.bar(
                        inverter_summary,
                        x='Inverter File',
                        y='Mean Efficiency',
                        title='Rata-Rata Inverter Efficiency',
                        color='Mean Efficiency',
                        range_y=[0,1]  # asumsikan efisiensi max = 1 (100%)
                    )
                    st.plotly_chart(fig_eff, use_container_width=True)
            else:
                st.warning("Tidak ada file inverter yang diupload, analisis inverter dilewati.")
        else:
            st.error("Mohon upload file EM & RM terlebih dahulu.")

if __name__ == "__main__":
    main()