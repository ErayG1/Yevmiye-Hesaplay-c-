import streamlit as st
import pandas as pd
import hashlib
import io

st.set_page_config(page_title="Aylik Yevmiye Hesaplayici", layout="wide")

st.title("📊 Aylik Yevmiye & Puantaj Hesaplayici")
st.write("Mobil veya bilgisayardan Excel dosyalarinizi yukleyerek 30 gunluk bir ozet raporu olusturabilirsiniz.")

if "uploaded_hashes" not in st.session_state:
    st.session_state.uploaded_hashes = set()

def get_file_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()

def read_excel_smart(file_bytes, file_name):
    """Excel, HTML tablosu veya CSV tabanli .xls dosyalarini akilli sekilde okur."""
    buffer = io.BytesIO(file_bytes)
    
    # 1. Deneme: Standart openpyxl veya xlrd ile okuma
    try:
        return pd.read_excel(buffer)
    except Exception:
        pass

    # 2. Deneme: Motoru xlrd olarak zorlayarak okuma (.xls)
    try:
        buffer.seek(0)
        return pd.read_excel(buffer, engine='xlrd')
    except Exception:
        pass

    # 3. Deneme: Bazi sistemler HTML tablosunu .xls uzantisiyla kaydeder
    try:
        buffer.seek(0)
        dfs = pd.read_html(buffer)
        if dfs:
            return dfs[0]
    except Exception:
        pass

    # 4. Deneme: CSV formatinda olma ihtimali
    try:
        buffer.seek(0)
        return pd.read_csv(buffer, sep=None, engine='python')
    except Exception:
        pass

    raise Exception("Excel dosya formati okunamadi. Lutfen dosya formatini kontrol edin.")

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
            df = read_excel_smart(file_bytes, file.name)
            
            # Sutun isimlerini standartlastirma
            df.columns = [str(col).strip().upper() for col in df.columns]

            # Alternatif sutun isimleri kontrolu
            column_mapping = {
                "TC NO": "TC",
                "TCKNO": "TC",
                "TC KİMLİK": "TC",
                "TC KIMLIK": "TC",
                "AD SOYAD": "ISIM",
                "ISIM SOYISIM": "ISIM",
                "AD SOYADI": "ISIM",
                "YEVMİYE": "YEVMIYE",
                "GUNLUK YEVMİYE": "YEVMIYE",
                "YEVMİYE TUTARI": "YEVMIYE"
            }
            df = df.rename(columns=column_mapping)

            if not {"TC", "ISIM", "YEVMIYE"}.issubset(df.columns):
                st.error(f"❌ **{file.name}** dosyasinda gerekli sutunlar (TC, ISIM, YEVMIYE) bulunamadi. Mevcut sutunlar: {list(df.columns)}")
                continue

            # Veri tiplerini duzenleme
            df["TC"] = df["TC"].astype(str).str.replace(".0", "", regex=False).str.strip()
            df["ISIM"] = df["ISIM"].astype(str).str.strip().str.title()
            df["YEVMIYE"] = pd.to_numeric(df["YEVMIYE"], errors="coerce").fillna(0)

            all_data.append(df[["TC", "ISIM", "YEVMIYE"]])
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
