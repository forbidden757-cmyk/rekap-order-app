import streamlit as st
import pymongo
import pandas as pd
from bson.objectid import ObjectId
from datetime import datetime
import json
import os
import math
from io import BytesIO
from xhtml2pdf import pisa
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

# Fungsi super kebal untuk menangani data kosong/NaN dari database lawas
def format_rupiah(angka):
    try:
        if pd.isna(angka): return "0"
        val = float(angka)
        if math.isnan(val): return "0"
        return f"{int(val):,}".replace(',', '.')
    except Exception:
        return "0"

def terbilang(n):
    try:
        if pd.isna(n): return ""
        val = float(n)
        if math.isnan(val): return ""
        n = int(val)
    except Exception: 
        return ""
        
    angka = ["", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas"]
    if n < 12: return angka[n]
    elif n < 20: return terbilang(n - 10) + " Belas"
    elif n < 100: return terbilang(n // 10) + " Puluh " + terbilang(n % 10)
    elif n < 200: return "Seratus " + terbilang(n - 100)
    elif n < 1000: return terbilang(n // 100) + " Ratus " + terbilang(n % 100)
    elif n < 2000: return "Seribu " + terbilang(n - 1000)
    elif n < 1000000: return terbilang(n // 1000) + " Ribu " + terbilang(n % 1000)
    elif n < 1000000000: return terbilang(n // 1000000) + " Juta " + terbilang(n % 1000000)
    else: return "Angka terlalu besar"

# ==========================================
# KONEKSI MONGODB ATLAS (Versi Aman untuk Cloud)
# ==========================================
# =========================================================================
# TODO: GANTI TEKS DI BAWAH INI DENGAN CONNECTION STRING DARI ATLAS ANDA!
# =========================================================================
ATLAS_URI = "mongodb+srv://db_free:Desindo20@cluster0.xrxgp9b.mongodb.net/?appName=Cluster0"

if 'mongo_client' not in st.session_state:
    try:
        # Menambahkan certifi agar lolos SSL di Streamlit Cloud
        st.session_state.mongo_client = pymongo.MongoClient(ATLAS_URI, serverSelectionTimeoutMS=5000, tlsCAFile=certifi.where())
        st.session_state.mongo_client.server_info() # Trigger koneksi
    except Exception as e:
        st.error(f"Gagal terhubung ke MongoDB Atlas. Detail: {e}")
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
                data = {
                    "kantor": input_kantor, "provinsi": input_provinsi,
                    "nama": input_nama, "jabatan": input_jabatan,
                    "npwp": input_npwp, "alamat": input_alamat
                }
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
            # Atribut key dihapus agar tidak error saat reload
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
                    data_order = {
                        "no_order": no_order, "tanggal": tgl_order.strftime("%Y-%m-%d"),
                        "id_pemesan": id_pemesan, "deskripsi": deskripsi,
                        "kuantitas": qty, "satuan": satuan,
                        "harga": harga, "total": total_nilai, "keterangan": keterangan
                    }
                    db.item_order.insert_one(data_order)
                    st.success("Item Order berhasil disimpan!")
                    st.rerun()

    st.markdown("---")
    st.subheader("Daftar Item Order")
    orders = list(db.item_order.find())
    if orders:
        df_order = pd.DataFrame(orders)
        df_order['_id'] = df_order['_id'].astype(str)
        df_order['id_pemesan'] = df_order['id_pemesan'].astype(str)
        
        # Mencegah error jika data lama tidak punya kolom harga/total
        if 'harga' not in df_order.columns: df_order['harga'] = 0
        if 'total' not in df_order.columns: df_order['total'] = 0
        
        df_order['harga_tampil'] = df_order['harga'].apply(lambda x: f"Rp {format_rupiah(x)}")
        df_order['total_tampil'] = df_order['total'].apply(lambda x: f"Rp {format_rupiah(x)}")
        
        st.dataframe(df_order[['no_order', 'tanggal', 'deskripsi', 'kuantitas', 'satuan', 'harga_tampil', 'total_tampil']], use_container_width=True, hide_index=True)
        
        col_del1, col_del2 = st.columns([3, 1])
        with col_del1:
            # Atribut key dihapus agar tidak error saat reload
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
    
    def get_auto_invoice():
        romawi = {1: 'I', 2: 'II', 3: 'III', 4: 'IV', 5: 'V', 6: 'VI', 7: 'VII', 8: 'VIII', 9: 'IX', 10: 'X', 11: 'XI', 12: 'XII'}
        now = datetime.now()
        prefix = f"KBM/LNU/{romawi[now.month]}/ELG-"
        last_inv = list(db.invoice.find({"no_invoice": {"$regex": f"^{prefix}"}}).sort("no_invoice", -1).limit(1))
        if last_inv:
            try: new_seq = int(last_inv[0]['no_invoice'].split('-')[-1]) + 1
            except ValueError: new_seq = 1
            return f"{prefix}{new_seq:03d}"
        return f"{prefix}001"

    # Hanya ambil order yang valid
    unik_orders = [o for o in db.item_order.distinct("no_order") if o]
    
    if not unik_orders:
        st.warning("⚠️ Belum ada Item Order. Silakan buat Order di Tab 2 dahulu.")
    else:
        col1, col2 = st.columns(2)
        with col1:
            inv_no = st.text_input("No Invoice", value=get_auto_invoice())
            inv_tgl = st.date_input("Tanggal Invoice")
            inv_order = st.selectbox("Pilih No Order yang Ditagihkan", unik_orders)
        with col2:
            inv_status = st.selectbox("Status Tagihan", ["Pending", "Dikirim", "Sudah Dibayar"])
            if inv_status == "Sudah Dibayar":
                inv_tgl_bayar = st.date_input("Tanggal Dibayar")
                inv_rek = st.selectbox("Rekening Tujuan", profil_pt.get("rekening", ["-"]))
            else:
                inv_tgl_bayar = "-"
                inv_rek = "-"
                
        if st.button("Buat Invoice", type="primary", use_container_width=True):
            data_inv = {
                "no_invoice": inv_no, "tanggal": inv_tgl.strftime("%Y-%m-%d"),
                "no_order": inv_order, "status": inv_status,
                "is_lunas": "Ya" if inv_status == "Sudah Dibayar" else "Tidak",
                "tgl_bayar": inv_tgl_bayar.strftime("%Y-%m-%d") if inv_status == "Sudah Dibayar" else "-",
                "rek_tujuan": inv_rek
            }
            try:
                db.invoice.insert_one(data_inv)
                st.success(f"Invoice {inv_no} berhasil diterbitkan!")
                st.rerun()
            except pymongo.errors.DuplicateKeyError:
                st.error("Gagal: No Invoice tersebut sudah pernah digunakan!")

    st.markdown("---")
    st.subheader("Daftar Invoice & Cetak PDF")
    invoices = list(db.invoice.find())
    
    if invoices:
        df_inv = pd.DataFrame(invoices)
        st.dataframe(df_inv[['no_invoice', 'tanggal', 'no_order', 'status', 'is_lunas', 'tgl_bayar', 'rek_tujuan']], use_container_width=True, hide_index=True)
        
        st.markdown("#### Cetak File PDF")
        pilih_cetak = st.selectbox("Pilih Invoice yang akan dicetak:", [inv.get('no_invoice', '') for inv in invoices])
        
        if pilih_cetak:
            inv_data = db.invoice.find_one({"no_invoice": pilih_cetak})
            items_db = list(db.item_order.find({"no_order": inv_data.get("no_order", "")}))
            
            if items_db:
                # Proteksi jika pemesan terhapus
                pemesan_data = db.pemesan.find_one({"_id": items_db[0].get("id_pemesan")})
                if not pemesan_data: pemesan_data = {}
                
                grand_total = sum(item.get("total", 0) for item in items_db)
                rekening_html = "<br>".join(profil_pt.get('rekening', []))
                
                tabel_item_html = ""
                for index, item in enumerate(items_db, start=1):
                    tabel_item_html += f"""
                    <tr>
                        <td align="center" style="padding: 8px; border-bottom: 1px solid #eee;">{index}</td>
                        <td style="padding: 8px; border-bottom: 1px solid #eee;">{item.get('deskripsi', '-')}</td>
                        <td align="center" style="padding: 8px; border-bottom: 1px solid #eee;">{item.get('kuantitas', '0')} {item.get('satuan', '')}</td>
                        <td align="right" style="padding: 8px; border-bottom: 1px solid #eee;">Rp {format_rupiah(item.get('harga', 0))}</td>
                        <td align="right" style="padding: 8px; border-bottom: 1px solid #eee; font-weight: bold;">Rp {format_rupiah(item.get('total', 0))}</td>
                    </tr>
                    """

                html_content = f"""
                <html>
                <head>
                    <style>
                        @page {{ size: a4 portrait; margin: 2cm; }}
                        body {{ font-family: Helvetica, sans-serif; font-size: 12px; color: #333; }}
                        .header {{ width: 100%; border-bottom: 2px solid #ccc; padding-bottom: 10px; margin-bottom: 20px; }}
                        .title {{ font-size: 24px; font-weight: bold; color: #2c3e50; }}
                        .table-items {{ width: 100%; border-collapse: collapse; margin-top: 20px; }}
                        .table-items th {{ background-color: #f8f9fa; padding: 10px; text-align: left; border-bottom: 2px solid #ddd; }}
                        .total-row {{ font-size: 14px; font-weight: bold; color: #e74c3c; }}
                        .terbilang {{ background-color: #f8f9fa; padding: 10px; border-left: 4px solid #3498db; margin-top: 20px; }}
                    </style>
                </head>
                <body>
                    <table class="header">
                        <tr>
                            <td width="50%"><div class="title">{profil_pt.get('nama_perusahaan', 'COMPANY NAME')}</div></td>
                            <td width="50%" align="right">
                                <strong style="font-size: 20px;">INVOICE</strong><br>
                                No: {inv_data.get('no_invoice', '')}<br>
                                Tanggal: {inv_data.get('tanggal', '')}
                            </td>
                        </tr>
                    </table>
                    
                    <table style="width: 100%; margin-bottom: 20px;">
                        <tr>
                            <td width="50%" valign="top">
                                <strong>DITAGIHKAN KEPADA:</strong><br>
                                {pemesan_data.get('nama','-')}<br>
                                {pemesan_data.get('kantor','-')}<br>
                                NPWP: {pemesan_data.get('npwp','-')}
                            </td>
                            <td width="50%" align="right" valign="top">
                                <strong>No Pesanan:</strong> {inv_data.get('no_order', '')}<br>
                                <strong>Status:</strong> {inv_data.get('status', '')}
                            </td>
                        </tr>
                    </table>

                    <table class="table-items">
                        <thead>
                            <tr>
                                <th width="5%" style="text-align: center;">No</th>
                                <th width="45%">Deskripsi</th>
                                <th width="10%" style="text-align: center;">Qty</th>
                                <th width="20%" style="text-align: right;">Harga Satuan</th>
                                <th width="20%" style="text-align: right;">Total</th>
                            </tr>
                        </thead>
                        <tbody>
                            {tabel_item_html}
                            <tr>
                                <td colspan="4" align="right" style="padding-top: 15px;"><strong>GRAND TOTAL</strong></td>
                                <td align="right" class="total-row" style="padding-top: 15px;">Rp {format_rupiah(grand_total)}</td>
                            </tr>
                        </tbody>
                    </table>

                    <div class="terbilang">
                        <strong>Terbilang:</strong><br>
                        <i>"{terbilang(grand_total)} Rupiah"</i>
                    </div>

                    <table style="width: 100%; margin-top: 40px;">
                        <tr>
                            <td width="60%" valign="top">
                                <strong>Instruksi Pembayaran:</strong><br>
                                {rekening_html}
                            </td>
                            <td width="40%" align="center" valign="bottom">
                                Dibuat Oleh,<br><br><br><br>
                                <strong><u>{profil_pt.get('pembuat_invoice', '_________________')}</u></strong><br>
                                {profil_pt.get('jabatan_pembuat', '')}
                            </td>
                        </tr>
                    </table>
                </body>
                </html>
                """

                pdf_buffer = BytesIO()
                pisa.CreatePDF(html_content, dest=pdf_buffer)
                
                st.download_button(
                    label=f"⬇️ Download Invoice {pilih_cetak} (PDF)",
                    data=pdf_buffer.getvalue(),
                    file_name=f"Invoice_{pilih_cetak.replace('/', '_')}.pdf",
                    mime="application/pdf",
                    type="primary"
                )
            else:
                st.warning("⚠️ Tidak bisa mencetak PDF. Item Order untuk invoice ini kosong atau sudah terhapus.")
