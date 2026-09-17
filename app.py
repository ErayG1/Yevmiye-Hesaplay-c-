import streamlit as st
import pandas as pd
import hashlib
import io
import re

st.set_page_config(page_title="Aylik Yevmiye Hesaplayici", layout="wide")

st.title("📊 Aylik Yevmiye & Puantaj Hesaplayici")
st.write("Mobil veya bilgisayardan Excel / Puantaj dosyalarinizi yukleyerek 30 gunluk ozet rapor olusturabilirsiniz.")

if "uploaded_hashes" not in st.session_state:
    st.session_state.uploaded_hashes = set()

def get_file_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

def clean_yevmiye_val(val):
    """'1 Yevmiye', '0.5 Yevmiye', '1,5' gibi metin ifadelerini sayiya cevirir."""
    if pd.isna(val):
        return 0.0
    val_str = str(val).replace(',', '.').strip()
    match = re.search(r"[-+]?\d*\.\d+|\d+", val_str)
    if match:
        return float(match.group())
    return 0.0

def extract_yevmiye_table(file_bytes, file_name):
    """Excel veya HTML bazli dosyalardan yevmiye tablosunu akilli sekilde ayiklar."""
    buffer = io.BytesIO(file_bytes)
    tables = []

    # 1. Deneme: HTML tablolari olarak okuma (Puantaj Raporu bu formattadir)
    try:
        buffer.seek(0)
        dfs = pd.read_html(buffer)
        tables.extend(dfs)
    except Exception:
        pass

    # 2. Deneme: Standart Excel okuma
    if not tables:
        try:
            buffer.seek(0)
            tables.append(pd.read_excel(buffer))
        except Exception:
            pass

    # 3. Deneme: xlrd motoru ile Excel okuma
    if not tables:
        try:
            buffer.seek(0)
            tables.append(pd.read_excel(buffer, engine='xlrd'))
        except Exception:
            pass

    # 4. Deneme: CSV okuma
    if not tables:
        try:
            buffer.seek(0)
            tables.append(pd.read_csv(buffer, sep=None, engine='python'))
        except Exception:
            pass

    if not tables:
        raise Exception("Dosya formati okunamadi.")

    column_mapping = {
        "TC NO": "TC", "TCKNO": "TC", "TC KİMLİK": "TC", "TC KIMLIK": "TC",
        "AD SOYAD": "ISIM", "ISIM SOYISIM": "ISIM", "AD SOYADI": "ISIM",
        "YEVMİYE": "YEVMIYE", "GUNLUK YEVMİYE": "YEVMIYE", "TOPLAM YEVMİYE": "YEVMIYE", "TOPLAM YEVMIYE": "YEVMIYE"
    }

    # Dosyadaki tum tablolar icinde arama yapip asil yevmiye tablosunu bulma
    for df in tables:
        df_cols = [str(col).strip().upper() for col in df.columns]
        df.columns = df_cols
        df_renamed = df.rename(columns=column_mapping)

        # TC, ISIM ve YEVMIYE sutunlari mevcut mu?
        if {"TC", "ISIM", "YEVMIYE"}.issubset(df_renamed.columns):
            return df_renamed[["TC", "ISIM", "YEVMIYE"]]

    # Eger sutun basligi 'TC' degil ama içerikte TC numaralari varsa ilk anlamli tabloyu bulma
    raise Exception(f"Dosyada (TC, ISIM, YEVMIYE) sutunlari iceren tablo bulunamadi.")

uploaded_files = st.file_uploader(
    "Yevmiye Excel Dosyalarini Secin (Coklu Secim)", 
    type=["xlsx", "xls", "csv"], 
    accept_multiple_files=True
)

if uploaded_files:
    all_data = []
    processed_count = 0
    duplicate_count = 0

    for file in uploaded_files:
        file_bytes = file.read()
        file_hash = get_file_hash(file_bytes)

        if file_hash in st.session_state.uploaded_hashes:
            st.warning(f"⚠️ **{file.name}** dosyasi daha once yuklendigi icin atlandi (Mukerrer kayit).")
            duplicate_count += 1
            continue

        try:
            df = extract_yevmiye_table(file_bytes, file.name)
            
            # Veri tiplerini ve formatlari temizleme
            df["TC"] = df["TC"].astype(str).str.replace(r"\.0$", "", regex=True).str.strip()
            df["ISIM"] = df["ISIM"].astype(str).str.strip().str.title()
            df["YEVMIYE"] = df["YEVMIYE"].apply(clean_yevmiye_val)

            all_data.append(df)
            st.session_state.uploaded_hashes.add(file_hash)
            processed_count += 1

        except Exception as e:
            st.error(f"❌ **{file.name}** okunurken hata olustu: {e}")

    if all_data:
        combined_df = pd.concat(all_data, ignore_index=True)

        summary_df = combined_df.groupby(["TC", "ISIM"]).agg(
            CALISILAN_GUN=("YEVMIYE", "count"),
            TOPLAM_YEVMIYE=("YEVMIYE", "sum"),
            ORTALAMA_YEVMIYE=("YEVMIYE", "mean")
        ).reset_index()

        summary_df["TOPLAM_YEVMIYE"] = summary_df["TOPLAM_YEVMIYE"].round(2)
        summary_df["ORTALAMA_YEVMIYE"] = summary_df["ORTALAMA_YEVMIYE"].round(2)

        st.success(f"✅ Toplam {processed_count} dosya basariyla islendi. ({duplicate_count} mukerrer dosya atlandi)")

        st.subheader("📋 30 Gunluk / Aylik Ozet Tablo")
        st.dataframe(summary_df, use_container_width=True)

        # Excel Indirme Butonu
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
            summary_df.to_excel(writer, index=False, sheet_name="Aylik Ozet")
        
        st.download_button(
            label="📥 Ozet Raporu Excel Olarak Indir",
            data=output.getvalue(),
            file_name="Aylik_Yevmiye_Ozet_Raporu.xlsx",
            mime="application/vnd.openpyxlformat-officedocument.spreadsheetml.sheet"
        )

if st.button("Sistemi ve Gecmisi Sifirla"):
    st.session_state.uploaded_hashes = set()
    st.rerun()
