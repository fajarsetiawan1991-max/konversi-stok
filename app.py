import re
import io

import pandas as pd
import streamlit as st

# ---------------------------------------------------------------------------
# Logika parsing (sama persis dengan versi Colab yang sudah teruji)
# ---------------------------------------------------------------------------
ROW = re.compile(r'^(\d+)\s+(.+?)\s+([A-Z]+\.[A-Z]+\.[A-Z]+)\s+(\d+\.\d+\.\d+)\s+(.*)$')


def to_pcs(qty, konv):
    """Contoh: 4.004.010 dengan konversi 72.12.1 -> 4*72 + 4*12 + 10 = 346 pcs"""
    return sum(int(a) * int(b) for a, b in zip(qty.split('.'), konv.split('.')))


def parse(text):
    data = []
    for line in text.splitlines():
        m = ROW.match(line)
        if m:
            kode, nama, satuan, konv, sisa = m.groups()
            t = sisa.split()
            data.append({
                'Kode Barang': kode,
                'Nama Barang': nama,
                'Satuan': satuan,
                'Konversi': konv,
                'Stok Fisik': t[-2],
                'Total PCS': to_pcs(t[-2], konv),
                'Amount (Rp)': int(t[-1].replace(',', '')),
            })
        # Sambungan nama barang di baris bawah (termasuk yang terpotong ganti halaman)
        elif data and not line[:6].strip() and line[6:32].strip() \
                and not line[32:].strip() and ':' not in line:
            data[-1]['Nama Barang'] += line[6:32].rstrip()
    return pd.DataFrame(data)


def to_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Stok')
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Tampilan
# ---------------------------------------------------------------------------
st.set_page_config(page_title="Konversi Stok Martupa", page_icon="📦")

st.title("📦 Konversi Stok Martupa")
st.caption("Upload file spool stok (.txt), lalu unduh hasilnya dalam format Excel.")

uploaded = st.file_uploader(
    "Pilih file .txt stok (bisa lebih dari satu)",
    type=["txt"],
    accept_multiple_files=True,
)

if uploaded:
    for f in uploaded:
        raw = f.read()
        try:
            text = raw.decode("latin-1")
        except Exception:
            text = raw.decode("utf-8", errors="ignore")

        df = parse(text)

        st.divider()
        st.subheader(f"📄 {f.name}")

        if df.empty:
            st.error("Format tidak dikenali — pastikan ini file laporan stok yang benar.")
            continue

        col1, col2 = st.columns(2)
        col1.metric("Jumlah Barang", f"{len(df)}")
        col2.metric("Total Amount", f"Rp {df['Amount (Rp)'].sum():,}")

        st.dataframe(df, use_container_width=True, hide_index=True)

        out_name = re.sub(r'\.txt$', '', f.name, flags=re.I) + '_HASIL.xlsx'
        st.download_button(
            label=f"⬇️ Unduh {out_name}",
            data=to_excel_bytes(df),
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
else:
    st.info("Belum ada file. Klik area di atas untuk mengunggah.")
