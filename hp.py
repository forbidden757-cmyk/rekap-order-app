import streamlit as st
import pymongo
import pandas as pd
from bson.objectid import ObjectId
from datetime import datetime
import json
import os
import math
from io import BytesIO
from fpdf import FPDF # Menggunakan FPDF2 yang lebih stabil di Cloud
import certifi

# --- KONFIGURASI HALAMAN ---
st.set_page_config(page_title="Sistem Rekap Order", page_icon="📦", layout="wide")

# ==========================================
# FUNGSI HELPER & PROFIL PERUSAHAAN
# ==========================================
PROFILE_FILE = 'company_profile.json'

def load_profile():
    if not os.path.exists(PROFILE_FILE):
        default_profile = {
            "nama_perusahaan": "PT. Perusahaan Maju Jaya",
            "deskripsi": "Penyedia Layanan Terbaik",
            "alamat": "Jl. Sudirman No. 1, Jakarta",
            "logo_path": "",
            "pembuat_invoice": "Joyboy",
            "jabatan_pembuat": "CEO",
            "ttd_path": "",
            "rekening": ["BCA - 1234567890 a/n PT Maju", "Mandiri - 0987654321 a/n PT Maju"]
        }
        with open(PROFILE_FILE, 'w') as f:
            json.dump(default_profile, f, indent=4)
    with open(PROFILE_FILE, 'r') as f:
        return json.load(f)

def format_rupiah(angka):
    try:
        if pd.isna(angka): return "0"
        val = float(angka)
        if math.isnan(val): return "0"
        return f"{int(val):,}".replace(',', '.')
    except Exception:
        return "0"

# ==========================================
# KONEKSI MONGODB ATLAS
# ==========================================
# GANTI LINK DI BAWAH INI DENGAN LINK ATLAS ANDA
ATLAS_URI = "mongodb+srv://db_free:Desindo20@cluster0.xrxgp9b.mongodb.net/?appName=Cluster0"

if 'mongo_client' not in st.session_state:
    try:
        # Kita gunakan opsi paling lengkap untuk menembus handshake TLS
        st.session_state.mongo_client = pymongo.MongoClient(
            ATLAS_URI,
            serverSelectionTimeoutMS=5000,
            tls=True,
            tlsAllowInvalidCertificates=True,
            connectTimeoutMS=30000,
            socketTimeoutMS=30000
        )
        # Tes koneksi dengan perintah ping sederhana
        st.session_state.mongo_client.admin.command('ping')
    except Exception as e:
        # Jika masih gagal, tampilkan pesan bantuan yang lebih jelas
        st.error(f"Gagal terhubung ke MongoDB Atlas.")
        st.info("Tips: Pastikan Password Anda tidak mengandung simbol seperti @, :, / atau #. Jika ada, ganti password di MongoDB Atlas dulu.")
        st.code(str(e))
        st.stop()

client = st.session_state.mongo_client
db = client["rekap_order_db"]
profil_pt = load_profile()

# --- HEADER APLIKASI ---
st.title("📦 Sistem Rekap Order & Invoice")
st.markdown(f"**{profil_pt.get('nama_perusahaan', 'Perusahaan')}** | Web Dashboard Internal")
st.markdown("---")

# --- MENU TABS ---
tab_pemesan, tab_order, tab_invoice = st.tabs(["👥 1. Data Pemesan", "🛒 2. Item Order", "🧾 3. Proses Invoice"])

# ==========================================
# TAB 1: DATA PEMESAN
# ==========================================
with tab_pemesan:
    st.subheader("Input Data Pemesan Baru")
    with st.form("form_tambah_pemesan", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            input_kantor = st.text_input("Kantor")
            input_nama = st.text_input("Nama Pemesan * (Wajib)")
            input_npwp = st.text_input("NPWP")
        with col2:
            input_provinsi = st.text_input("Provinsi")
            input_jabatan = st.text_input("Jabatan")
            input_alamat = st.text_area("Alamat Lengkap")
            
        btn_simpan_pemesan = st.form_submit_button("Simpan Pemesan", type="primary")
        
        if btn_simpan_pemesan:
            if not input_nama:
                st.error("Gagal: Nama Pemesan tidak boleh kosong!")
            else:
                data = {"kantor": input_kantor, "provinsi": input_provinsi, "nama": input_nama, "jabatan": input_jabatan, "npwp": input_npwp, "alamat": input_alamat}
                db.pemesan.insert_one(data)
                st.success("Data berhasil disimpan!")
                st.rerun()

    st.markdown("---")
    st.subheader("Daftar Pemesan")
    pemesans = list(db.pemesan.find())
    if pemesans:
        df_pemesan = pd.DataFrame(pemesans)
        df_pemesan['_id'] = df_pemesan['_id'].astype(str)
        st.dataframe(df_pemesan[['_id', 'nama', 'kantor', 'jabatan', 'provinsi', 'npwp', 'alamat']], use_container_width=True, hide_index=True)
        
        col_del1, col_del2 = st.columns([3, 1])
        with col_del1:
            pilihan_hapus_pemesan = st.selectbox("Pilih Pemesan yang ingin dihapus:", [f"{p['_id']} - {p.get('nama', '-')}" for p in pemesans])
        with col_del2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Hapus Pemesan"):
                db.pemesan.delete_one({"_id": ObjectId(pilihan_hapus_pemesan.split(" - ")[0])})
                st.success("Data berhasil dihapus!")
                st.rerun()

# ==========================================
# TAB 2: ITEM ORDER
# ==========================================
with tab_order:
    st.subheader("Input Item Order Baru")
    pemesans = list(db.pemesan.find())
    
    if not pemesans:
        st.warning("⚠️ Silakan isi Data Pemesan di Tab 1 terlebih dahulu!")
    else:
        with st.form("form_tambah_order", clear_on_submit=True):
            col1, col2, col3 = st.columns(3)
            with col1:
                no_order = st.text_input("No Order * (Wajib)")
                tgl_order = st.date_input("Tanggal Order")
                pilihan_pemesan = st.selectbox("Pilih Pemesan", [f"{p['_id']} - {p.get('nama', '-')} ({p.get('kantor', '-')})" for p in pemesans])
            with col2:
                deskripsi = st.text_input("Deskripsi Barang/Jasa")
                qty = st.number_input("Kuantitas (Qty)", min_value=1, value=1, step=1)
                satuan = st.text_input("Satuan (Pcs/Box/Paket)")
            with col3:
                harga = st.number_input("Harga Satuan (Rp)", min_value=0, value=0, step=1000)
                keterangan = st.text_area("Keterangan Tambahan")
                
            btn_simpan_order = st.form_submit_button("Simpan Item Order", type="primary")
            
            if btn_simpan_order:
                if not no_order:
                    st.error("No Order wajib diisi!")
                else:
                    id_pemesan = ObjectId(pilihan_pemesan.split(" - ")[0])
                    total_nilai = qty * harga
                    data_order = {"no_order": no_order, "tanggal": tgl_order.strftime("%Y-%m-%d"), "id_pemesan": id_pemesan, "deskripsi": deskripsi, "kuantitas": qty, "satuan": satuan, "harga": harga, "total": total_nilai, "keterangan": keterangan}
                    db.item_order.insert_one(data_order)
                    st.success("Item Order berhasil disimpan!")
                    st.rerun()

    st.markdown("---")
    st.subheader("Daftar Item Order")
    orders = list(db.item_order.find())
    if orders:
        df_order = pd.DataFrame(orders)
        df_order['_id'] = df_order['_id'].astype(str)
        df_order['harga_tampil'] = df_order['harga'].apply(lambda x: f"Rp {format_rupiah(x)}")
        df_order['total_tampil'] = df_order['total'].apply(lambda x: f"Rp {format_rupiah(x)}")
        st.dataframe(df_order[['no_order', 'tanggal', 'deskripsi', 'kuantitas', 'satuan', 'harga_tampil', 'total_tampil']], use_container_width=True, hide_index=True)
        
        col_del1, col_del2 = st.columns([3, 1])
        with col_del1:
            pilihan_hapus_order = st.selectbox("Pilih Item Order yang ingin dihapus:", [f"{o['_id']} - {o.get('no_order', '-')} ({o.get('deskripsi', '-')})" for o in orders])
        with col_del2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("Hapus Item"):
                db.item_order.delete_one({"_id": ObjectId(pilihan_hapus_order.split(" - ")[0])})
                st.success("Item berhasil dihapus!")
                st.rerun()

# ==========================================
# TAB 3: PROSES INVOICE
# ==========================================
with tab_invoice:
    st.subheader("Generate Invoice Baru")
    unik_orders = [o for o in db.item_order.distinct("no_order") if o]
    
    if not unik_orders:
        st.warning("⚠️ Belum ada Item Order. Silakan buat Order di Tab 2 dahulu.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            inv_no = st.text_input("No Invoice")
            inv_tgl = st.date_input("Tanggal Invoice")
            inv_order = st.selectbox("Pilih No Order yang Ditagihkan", unik_orders)
        with col2:
            inv_status = st.selectbox("Status Tagihan", ["Pending", "Dikirim", "Sudah Dibayar"])
            inv_rek = st.selectbox("Rekening Tujuan", profil_pt.get("rekening", ["-"]))
                
        if st.button("Buat Invoice", type="primary", use_container_width=True):
            data_inv = {"no_invoice": inv_no, "tanggal": inv_tgl.strftime("%Y-%m-%d"), "no_order": inv_order, "status": inv_status, "rek_tujuan": inv_rek}
            db.invoice.insert_one(data_inv)
            st.success(f"Invoice {inv_no} berhasil diterbitkan!")
            st.rerun()

    st.markdown("---")
    st.subheader("Daftar Invoice & Cetak PDF")
    invoices = list(db.invoice.find())
    
    if invoices:
        df_inv = pd.DataFrame(invoices)
        st.dataframe(df_inv[['no_invoice', 'tanggal', 'no_order', 'status', 'rek_tujuan']], use_container_width=True, hide_index=True)
        
        st.markdown("#### Cetak File PDF")
        pilih_cetak = st.selectbox("Pilih Invoice:", [inv.get('no_invoice', '') for inv in invoices])
        
        if pilih_cetak and st.button("Generate Download Link"):
            inv_data = db.invoice.find_one({"no_invoice": pilih_cetak})
            items_db = list(db.item_order.find({"no_order": inv_data.get("no_order", "")}))
            
            if items_db:
                pemesan_data = db.pemesan.find_one({"_id": items_db[0].get("id_pemesan")}) or {}
                grand_total = sum(item.get("total", 0) for item in items_db)

                # PROSES PDF DENGAN FPDF2
                pdf = FPDF()
                pdf.add_page()
                pdf.set_font("Helvetica", "B", 16)
                pdf.cell(0, 10, profil_pt.get('nama_perusahaan', 'INVOICE'), ln=True)
                pdf.set_font("Helvetica", "", 10)
                pdf.cell(0, 5, f"No: {inv_data.get('no_invoice')}", ln=True)
                pdf.cell(0, 5, f"Tanggal: {inv_data.get('tanggal')}", ln=True)
                pdf.ln(5)
                pdf.set_font("Helvetica", "B", 10)
                pdf.cell(0, 5, "KEPADA:", ln=True)
                pdf.set_font("Helvetica", "", 10)
                pdf.cell(0, 5, f"{pemesan_data.get('nama', '-')}", ln=True)
                pdf.cell(0, 5, f"{pemesan_data.get('kantor', '-')}", ln=True)
                pdf.ln(10)
                # Tabel
                pdf.set_fill_color(240, 240, 240)
                pdf.cell(100, 8, "Deskripsi", 1, 0, "L", True)
                pdf.cell(30, 8, "Qty", 1, 0, "C", True)
                pdf.cell(60, 8, "Total", 1, 1, "R", True)
                for item in items_db:
                    pdf.cell(100, 8, str(item.get('deskripsi', '-')), 1)
                    pdf.cell(30, 8, f"{item.get('kuantitas')} {item.get('satuan')}", 1, 0, "C")
                    pdf.cell(60, 8, f"Rp {format_rupiah(item.get('total', 0))}", 1, 1, "R")
                pdf.set_font("Helvetica", "B", 11)
                pdf.cell(130, 10, "GRAND TOTAL  ", 0, 0, "R")
                pdf.cell(60, 10, f"Rp {format_rupiah(grand_total)}", 0, 1, "R")

                pdf_output = pdf.output()
                st.download_button(label="⬇️ Download PDF", data=bytes(pdf_output), file_name=f"{pilih_cetak}.pdf", mime="application/pdf")
