import imaplib
import email
from email.header import decode_header
import html2text

# Server IMAP Yandex dan kredensial
imap_server = "imap.yandex.com"
alamat_email = "EMAIL"
kata_sandi = "PASSWORD_APP_IMAP"

# Terhubung ke server IMAP Yandex
try:
    imap = imaplib.IMAP4_SSL(imap_server)
    imap.login(alamat_email, kata_sandi)
    imap.select("INBOX")  # Pilih kotak surat yang ingin Anda ambil emailnya

    # Cari semua email dalam kotak surat
    status, email_ids = imap.search(None, "ALL")
    
    # Dapatkan daftar ID email
    daftar_id_email = email_ids[0].split()

    # Inisialisasi objek html2text
    html_converter = html2text.HTML2Text()

    # Ambil dan simpan setiap email dalam file teks
    for id_email in daftar_id_email:
        status, data_email = imap.fetch(id_email, "(RFC822)")
        email_raw = data_email[0][1]
        pesan = email.message_from_bytes(email_raw)
        
        # Ekstrak informasi email
        subjek, enkoding = decode_header(pesan["Subject"])[0]
        dari = pesan.get("From")
        tanggal = pesan.get("Date")
        
        print(f"Subjek: {subjek}")
        print(f"Dari: {dari}")
        print(f"Tanggal: {tanggal}")
        
        # Konversi isi email dari HTML ke teks
        isi_email_html = pesan.get_payload()
        isi_email_teks = html_converter.handle(isi_email_html)
        
        nama_file = f"{subjek}.txt"  # Gunakan subjek sebagai nama file
        with open(nama_file, "w", encoding="utf-8") as file:
            file.write(isi_email_teks)
        print(f"Isi Email disimpan dalam file: {nama_file}\n")
        
    # Logout dari server IMAP
    imap.logout()

except Exception as e:
    print(f"Terjadi kesalahan: {str(e)}")
