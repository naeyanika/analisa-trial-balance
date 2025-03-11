import streamlit as st
import pandas as pd
import numpy as np
import io
import matplotlib.pyplot as plt
import seaborn as sns

# Set page configuration
st.set_page_config(page_title="Financial Data Analysis", layout="wide")

# Custom CSS to improve appearance
st.markdown("""
<style>
    .highlight-green {
        background-color: rgba(0, 255, 0, 0.2);
    }
    .highlight-yellow {
        background-color: rgba(255, 255, 0, 0.2);
    }
    .highlight-red {
        background-color: rgba(255, 0, 0, 0.2);
    }
    .main-header {
        font-size: 2.5rem;
        font-weight: bold;
        margin-bottom: 1rem;
    }
    .sub-header {
        font-size: 1.5rem;
        font-weight: bold;
        margin-top: 1.5rem;
        margin-bottom: 0.5rem;
    }
</style>
""", unsafe_allow_html=True)

# Main title
st.markdown('<p class="main-header">Analisis Perubahan Bulanan Keuangan</p>', unsafe_allow_html=True)

# File upload
st.markdown('<p class="sub-header">Upload Data Excel</p>', unsafe_allow_html=True)
uploaded_file = st.file_uploader("Pilih file Excel yang berisi data keuangan", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        # Read the Excel file
        df = pd.read_excel(uploaded_file)
        
        # Check if the required columns exist
        required_columns = ["No Akun", "Keterangan"]
        if not all(col in df.columns for col in required_columns):
            st.error("File Excel harus memiliki kolom 'No Akun' dan 'Keterangan'")
        else:
            # Display the raw data
            st.markdown('<p class="sub-header">Data Mentah</p>', unsafe_allow_html=True)
            st.write(df)
            
            # Identify month columns
            month_columns = [col for col in df.columns if col not in required_columns]
            
            if len(month_columns) < 2:
                st.error("Data harus memiliki minimal 2 bulan untuk melakukan analisis perubahan")
            else:
                # Convert numeric columns to numeric
                for col in month_columns:
                    df[col] = pd.to_numeric(df[col], errors='coerce')
                
                # Function to calculate percentage change
                def calculate_change(row, current_month, previous_month):
                    if pd.isna(row[previous_month]) or pd.isna(row[current_month]):
                        return np.nan
                    if row[previous_month] == 0:
                        return np.inf if row[current_month] > 0 else 0
                    return (row[current_month] - row[previous_month]) / row[previous_month] * 100
                
                # Create a dataframe for monthly changes
                changes_df = df[["No Akun", "Keterangan"]].copy()
                
                # Calculate month-to-month changes in percentage
                for i in range(1, len(month_columns)):
                    current_month = month_columns[i]
                    previous_month = month_columns[i-1]
                    col_name = f"Perubahan {previous_month} ke {current_month} (%)"
                    changes_df[col_name] = df.apply(lambda row: calculate_change(row, current_month, previous_month), axis=1)
                
                # Calculate month-to-month changes in absolute value
                absolute_changes_df = df[["No Akun", "Keterangan"]].copy()
                for i in range(1, len(month_columns)):
                    current_month = month_columns[i]
                    previous_month = month_columns[i-1]
                    col_name = f"Perubahan {previous_month} ke {current_month} (Rp)"
                    absolute_changes_df[col_name] = df[current_month] - df[previous_month]
                
                # Function to color code changes
                def color_significant_changes(val):
                    if pd.isna(val):
                        return ''
                    if isinstance(val, (int, float)):
                        if -5 <= val <= 5:
                            return 'background-color: rgba(0, 255, 0, 0.2)'  # Green
                        elif -20 <= val <= 20:
                            return 'background-color: rgba(255, 255, 0, 0.2)'  # Yellow
                        else:
                            return 'background-color: rgba(255, 0, 0, 0.2)'  # Red
                    return ''
                
                # Analyze specific categories
                st.markdown('<p class="sub-header">Analisis Perubahan Bulanan</p>', unsafe_allow_html=True)
                
                # Filter expense categories
                expense_categories = [
                    "ATK", "Foto Copy", "Cetakan", 
                    "Telephone", 
                    "Komputer/IT", 
                    "BBM/Transport", 
                    "Transport Lainnya", 
                    "Listrik & Air", 
                    "Sewa", 
                    "Perlengkapan Kantor",
                    "Pengiriman", 
                    "Konsumsi", 
                    "Kantor Lainnya", 
                    "Perawatan Gedung", "Perawatan Kantor",
                    "Perawatan Kendaraan", 
                    "Perawatan Komputer/IT", 
                    "Penyusutan Kendaraan", 
                    "Penyusutan komputer/IT", 
                    "Peny Perl Elektronik", 
                    "Penyusutan peralatan kan", 
                    "administrasi Bank", 
                    "Elektronik", 
                    "Sumbangan", 
                    "Perijinan", 
                    "Kebersihan"
                ]
                
                expense_filter = df['Keterangan'].apply(lambda x: any(category.lower() in str(x).lower() for category in expense_categories))
                expense_df = df[expense_filter].copy()
                
                # Filter pinjaman rows
                pinjaman_categories = [
                    "Pinjaman Umum",
                    "Pinjaman Micro Bisnis",
                    "Pinjaman Renovasi Rumah",
                    "Pinjaman Pendidikan Anak",
                    "Pinjaman Sanitasi",
                    "Pinjaman Alat Rumah Tangga"
                ]
                
                pinjaman_filter = df['Keterangan'].apply(lambda x: x.strip().lower() in [category.lower() for category in pinjaman_categories])
                pinjaman_df = df[pinjaman_filter].copy()
                
                # Filter simpanan rows
                simpanan_categories = [
                    "Simpanan Sukarela",
                    "Simpanan Hari Raya",
                    "Simpanan Qurban",
                    "Simpanan Pokok",
                    "Simpanan Wajib",
                    "Simpanan Pendidikan & Kesehatan",
                    "Dana Pensiun Anggota"
                ]
                
                simpanan_filter = df['Keterangan'].apply(lambda x: x.strip().lower() in [category.lower() for category in simpanan_categories])
                simpanan_df = df[simpanan_filter].copy()
                
                # Show expense analysis
                if not expense_df.empty:
                    st.markdown("### Analisis Biaya")
                    styled_expense_changes = absolute_changes_df[absolute_changes_df['No Akun'].isin(expense_df['No Akun'])]
                    
                    # Apply styling based on percentage changes
                    pct_changes = changes_df[changes_df['No Akun'].isin(expense_df['No Akun'])]
                    pct_cols = [col for col in pct_changes.columns if "Perubahan" in col]
                    styled_pct_changes = pct_changes.style.applymap(color_significant_changes, subset=pct_cols)
                    
                    # Display both tables
                    st.write("Perubahan Persentase:")
                    st.dataframe(styled_pct_changes)
                    
                    st.write("Perubahan Nominal (Rp):")
                    st.dataframe(styled_expense_changes)
                
                # Show pinjaman analysis
                if not pinjaman_df.empty:
                    st.markdown("### Analisis Pinjaman")
                    styled_pinjaman_changes = absolute_changes_df[absolute_changes_df['No Akun'].isin(pinjaman_df['No Akun'])]
                    
                    # Apply styling based on percentage changes
                    pct_changes = changes_df[changes_df['No Akun'].isin(pinjaman_df['No Akun'])]
                    pct_cols = [col for col in pct_changes.columns if "Perubahan" in col]
                    styled_pct_changes = pct_changes.style.applymap(color_significant_changes, subset=pct_cols)

                    # Grouped bar chart for loan trends
                    st.markdown('<p class="sub-header">Tren Pinjaman</p>', unsafe_allow_html=True)
                    
                    # Prepare data for grouped bar chart
                    loan_data = pinjaman_df.set_index('Keterangan')[month_columns].T
                    loan_data.columns.name = 'Keterangan'
                    loan_data.index.name = 'Bulan, Tahun'
                    
                    # Plot grouped bar chart
                    x = np.arange(len(loan_data.index))
                    width = 0.25
                    multiplier = 0
                    fig, ax = plt.subplots(layout='constrained', figsize=(12, 8))
                    
                    for category, values in loan_data.items():
                        offset = width * multiplier
                        rects = ax.bar(x + offset, values / 1e8, width, label=category)  # Convert to ratusan juta
                        ax.bar_label(rects, padding=3)
                        multiplier += 1
                    
                    # Add some text for labels, title and custom x-axis tick labels, etc.
                    ax.set_ylabel('Nominal (ratusan juta Rp)')
                    ax.set_title('Tren Pinjaman Bulanan')
                    ax.set_xticks(x + width, loan_data.index)
                    ax.legend(loc='upper left', ncols=3)
                    ax.set_ylim(0, loan_data.max().max() / 1e8 + 1)
                    
                    plt.xticks(rotation=45)
                    st.pyplot(fig)
                    
                    
                    
                    # Display both tables
                    st.write("Perubahan Persentase:")
                    st.dataframe(styled_pct_changes)
                    
                    st.write("Perubahan Nominal (Rp):")
                    st.dataframe(styled_pinjaman_changes)
                    
                    
                # Show simpanan analysis
                if not simpanan_df.empty:
                    st.markdown("### Analisis Simpanan")
                    styled_simpanan_changes = absolute_changes_df[absolute_changes_df['No Akun'].isin(simpanan_df['No Akun'])]
                    
                    # Apply styling based on percentage changes
                    pct_changes = changes_df[changes_df['No Akun'].isin(simpanan_df['No Akun'])]
                    pct_cols = [col for col in pct_changes.columns if "Perubahan" in col]
                    styled_pct_changes = pct_changes.style.applymap(color_significant_changes, subset=pct_cols)
                    
                    # Display both tables
                    st.write("Perubahan Persentase:")
                    st.dataframe(styled_pct_changes)
                    
                    st.write("Perubahan Nominal (Rp):")
                    st.dataframe(styled_simpanan_changes)
                
                # Create visualizations
                st.markdown('<p class="sub-header">Visualisasi Data</p>', unsafe_allow_html=True)
                
                # Create monthly trend charts for selected categories
                categories_to_visualize = st.multiselect(
                    "Pilih kategori untuk divisualisasikan:", 
                    df["Keterangan"].unique()
                )
                
                if categories_to_visualize:
                    for category in categories_to_visualize:
                        category_data = df[df["Keterangan"] == category]
                        if not category_data.empty:
                            st.markdown(f"#### Tren Bulanan: {category}")
                            category_values = category_data[month_columns].values[0]
                            
                            fig, ax = plt.subplots(figsize=(12, 6))
                            ax.plot(month_columns, category_values, marker='o', linewidth=2)
                            ax.set_title(f"Tren Bulanan: {category}")
                            ax.set_ylabel("Nilai (Rp)")
                            ax.set_xlabel("Bulan")
                            ax.grid(True)
                            plt.xticks(rotation=45)
                            st.pyplot(fig)
                
                # Generate summary report
                st.markdown('<p class="sub-header">Laporan Ringkasan</p>', unsafe_allow_html=True)
                
                # Find significant changes
                def find_significant_changes(changes_df, category_filter, threshold=20):
                    filtered_df = changes_df[category_filter].copy()
                    change_cols = [col for col in filtered_df.columns if "Perubahan" in col]
                    
                    significant_changes = []
                    for idx, row in filtered_df.iterrows():
                        for col in change_cols:
                            if pd.notna(row[col]) and abs(row[col]) > threshold:
                                period = col.replace("Perubahan ", "").replace(" (%)", "")
                                significant_changes.append({
                                    "Kategori": row["Keterangan"],
                                    "No Akun": row["No Akun"],
                                    "Periode": period,
                                    "Perubahan (%)": row[col]
                                })
                    
                    return pd.DataFrame(significant_changes).sort_values("Perubahan (%)", ascending=False)
                
                # Generate summary of significant changes
                expense_significant = find_significant_changes(changes_df, expense_filter)
                pinjaman_significant = find_significant_changes(changes_df, pinjaman_filter)
                simpanan_significant = find_significant_changes(changes_df, simpanan_filter)
                
                # Create summary report
                summary_report = []
                
                if not expense_significant.empty:
                    top_expense = expense_significant.iloc[0]
                    summary_report.append(f"1. Perubahan biaya terbesar terjadi pada kategori '{top_expense['Kategori']}' pada periode {top_expense['Periode']} dengan perubahan {top_expense['Perubahan (%)']}%.")
                
                if not pinjaman_significant.empty:
                    top_pinjaman = pinjaman_significant.iloc[0]
                    summary_report.append(f"2. Pinjaman mengalami perubahan signifikan pada kategori '{top_pinjaman['Kategori']}' pada periode {top_pinjaman['Periode']} dengan perubahan {top_pinjaman['Perubahan (%)']}%.")
                
                if not simpanan_significant.empty:
                    top_simpanan = simpanan_significant.iloc[0]
                    summary_report.append(f"3. Simpanan mengalami perubahan signifikan pada kategori '{top_simpanan['Kategori']}' pada periode {top_simpanan['Periode']} dengan perubahan {top_simpanan['Perubahan (%)']}%.")
                
                st.markdown("#### Temuan Utama:")
                for finding in summary_report:
                    st.write(finding)
                
                # Generate customized analysis report
                st.markdown('<p class="sub-header">Laporan Analisis Keuangan</p>', unsafe_allow_html=True)
                
                report_text = st.text_area(
                    "Edit laporan analisis untuk di-download:",
                    f"""
Berdasarkan Trial Balance secara time series pada periode {month_columns[0]} s.d. {month_columns[-1]}, dapat dilihat beberapa anomali/ketidakwajaran dengan penjelasan singkat sebagai berikut:

1. Analisis Biaya Operasional:
   - Kategori biaya dengan perubahan terbesar adalah pada akun yang terkait dengan pengeluaran operasional.
   - Terdapat fluktuasi signifikan pada beberapa bulan tertentu yang perlu diperhatikan.

2. Analisis Pinjaman:
   - Pinjaman mengalami perubahan dinamis selama periode pengamatan.
   - Perlu perhatian khusus pada kenaikan/penurunan yang terjadi secara ekstrem.

3. Analisis Simpanan:
   - Simpanan juga menunjukkan tren perubahan yang perlu dimonitor.
   - Beberapa perubahan ekstrem dapat mengindikasikan pergerakan dana yang tidak biasa.

Rekomendasi:
- Lakukan pemeriksaan lebih lanjut terhadap perubahan ekstrem yang terjadi.
- Buat mekanisme monitoring yang lebih ketat untuk kategori biaya yang sering mengalami fluktuasi.
- Evaluasi kebijakan pinjaman dan simpanan untuk memastikan kestabilan keuangan.
""",
                    height=400
                )
                
                # Export to Excel
                st.markdown('<p class="sub-header">Download Hasil Analisis</p>', unsafe_allow_html=True)
                
                if st.button("Generate Excel Report"):
                    buffer = io.BytesIO()
                    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
                        # Write original data
                        df.to_excel(writer, sheet_name='Data Asli', index=False)
                        
                        # Write percentage changes
                        changes_df.to_excel(writer, sheet_name='Perubahan (%)', index=False)
                        
                        # Write absolute changes
                        absolute_changes_df.to_excel(writer, sheet_name='Perubahan (Rp)', index=False)
                        
                        # Write expense analysis
                        if not expense_df.empty:
                            expense_df.to_excel(writer, sheet_name='Analisis Biaya', index=False)
                            changes_df[changes_df['No Akun'].isin(expense_df['No Akun'])].to_excel(
                                writer, sheet_name='Perubahan Biaya (%)', index=False
                            )
                        
                        # Write pinjaman analysis
                        if not pinjaman_df.empty:
                            pinjaman_df.to_excel(writer, sheet_name='Analisis Pinjaman', index=False)
                            changes_df[changes_df['No Akun'].isin(pinjaman_df['No Akun'])].to_excel(
                                writer, sheet_name='Perubahan Pinjaman (%)', index=False
                            )
                        
                        # Write simpanan analysis
                        if not simpanan_df.empty:
                            simpanan_df.to_excel(writer, sheet_name='Analisis Simpanan', index=False)
                            changes_df[changes_df['No Akun'].isin(simpanan_df['No Akun'])].to_excel(
                                writer, sheet_name='Perubahan Simpanan (%)', index=False
                            )

                        # Show pinjaman analysis
                        if not pinjaman_df.empty:
                            st.markdown("### Analisis Pinjaman")
                            styled_pinjaman_changes = absolute_changes_df[absolute_changes_df['No Akun'].isin(pinjaman_df['No Akun'])]
    
                            # Apply styling based on percentage changes
                            pct_changes = changes_df[changes_df['No Akun'].isin(pinjaman_df['No Akun'])]
                            pct_cols = [col for col in pct_changes.columns if "Perubahan" in col]
                            styled_pct_changes = pct_changes.style.applymap(color_significant_changes, subset=pct_cols)

                            # Komposisi Pinjaman
                            st.markdown('<p class="sub-header">Komposisi Pinjaman</p>', unsafe_allow_html=True)
    
                            # Filter pinjaman untuk bulan dan tahun terakhir
                            last_month = month_columns[-1]
                            pinjaman_last_month = pinjaman_df[last_month].dropna()
    
                            # Hitung total nominal dan persentase komposisi
                            total_pinjaman = pinjaman_last_month.sum()
                            pinjaman_composition = pinjaman_last_month.groupby(pinjaman_df['Keterangan']).sum().reset_index()
                            pinjaman_composition['Persentase (%)'] = (pinjaman_composition[last_month] / total_pinjaman) * 100
    
                            # Tampilkan tabel komposisi pinjaman
                            st.markdown("#### Komposisi Pinjaman Bulan Terakhir")
                            st.dataframe(pinjaman_composition)
                            
                            # Buat pie chart
                            fig, ax = plt.subplots(figsize=(8, 8))

                            # Tampilkan semua label di luar pie chart
                            labels = pinjaman_composition['Persentase (%)'].apply(lambda x: f"{x:.1f}%")

                            wedges, texts, autotexts = ax.pie(
                                pinjaman_composition[last_month], 
                                labels=labels, 
                                autopct='%1.1f%%', 
                                startangle=90, 
                                colors=sns.color_palette("Set3", len(pinjaman_composition)),
                                labeldistance=1.1
                                )
                            
    
                        # Tambahkan legenda
                            ax.legend(wedges, pinjaman_composition['Keterangan'], title="Keterangan", loc="center left", bbox_to_anchor=(1, 0, 0.5, 1))
                            ax.axis('equal')  # Equal aspect ratio ensures that pie is drawn as a circle.

                            st.pyplot(fig)

                        # Write significant changes
                        pd.concat([
                            expense_significant, 
                            pinjaman_significant,
                            simpanan_significant
                        ]).to_excel(writer, sheet_name='Perubahan Signifikan', index=False)
                        
                        # Write summary report
                        workbook = writer.book
                        summary_sheet = workbook.add_worksheet('Ringkasan Analisis')
                        
                        # Format the summary sheet
                        bold_format = workbook.add_format({'bold': False, 'font_size': 11})
                        normal_format = workbook.add_format({'font_size': 11})
                        
                        # Write the report
                        summary_sheet.write(0, 0, "LAPORAN ANALISIS KEUANGAN", bold_format)
                        summary_sheet.write(2, 0, "Periode Analisis:", bold_format)
                        summary_sheet.write(2, 1, f"{month_columns[0]} s.d. {month_columns[-1]}", normal_format)
                        
                        row = 4
                        for line in report_text.split('\n'):
                            summary_sheet.write(row, 0, line, normal_format)
                            row += 1
                    
                    buffer.seek(0)
                    st.download_button(
                        label="Download Excel Report",
                        data=buffer,
                        file_name="Analisis_Keuangan.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )
                
    except Exception as e:
        st.error(f"Error reading file: {e}")
else:
    st.info("Silakan upload file Excel untuk memulai analisis")

# Add footer with instructions
st.markdown("---")
st.markdown("""
### Panduan Penggunaan:
1. Upload file Excel dengan format sesuai (kolom No Akun, Keterangan, dan kolom bulan-bulan)
2. Aplikasi akan menganalisis perubahan bulanan untuk kategori biaya, pinjaman, dan simpanan
3. Hasil analisis dapat didownload dalam format Excel
4. Warna pada tabel persentase perubahan:
   - Hijau: Perubahan ringan (-5% sampai 5%)
   - Kuning: Perubahan moderat (-20% sampai 20%)
   - Merah: Perubahan signifikan (di bawah -20% atau di atas 20%)
""")
