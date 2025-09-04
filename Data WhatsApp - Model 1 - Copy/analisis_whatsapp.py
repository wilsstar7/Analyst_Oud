import re
import os
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns
from wordcloud import WordCloud
import arabic_reshaper
from bidi.algorithm import get_display
from collections import Counter

def parse_whatsapp_chat(folder_path):
    """
    Membaca semua file .txt di sebuah folder, mem-parsingnya, 
    dan mengubahnya menjadi DataFrame Pandas.
    """
    all_data = []
    
    # Pola Regex ini disesuaikan untuk format WhatsApp: DD/MM/YY HH.MM - Sender: Message
    # Contoh: 20/08/25 02.54 - Nama Pengirim: Pesan
    # Dibuat lebih spesifik untuk format tahun 2 digit (YY) untuk meningkatkan akurasi.
    pattern = re.compile(r"(\d{1,2}/\d{1,2}/\d{2}) (\d{1,2}\.\d{2}) - (.*?): (.*)")

    for filename in os.listdir(folder_path):
        if filename.endswith(".txt"):
            file_path = os.path.join(folder_path, filename)
            with open(file_path, 'r', encoding='utf-8') as f:
                for line in f:
                    match = pattern.match(line)
                    if match:
                        date_str, time_str, sender, message = match.groups()
                        # Membersihkan karakter kontrol tak terlihat (seperti LTR/RTL marks) dari nama pengirim
                        cleaned_sender = ''.join(c for c in sender if c.isprintable()).strip()
                        all_data.append([f"{date_str}, {time_str}", cleaned_sender, message.strip(), filename])
    
    if not all_data:
        return pd.DataFrame() 

    df = pd.DataFrame(all_data, columns=['Timestamp', 'Sender', 'Message', 'Filename'])
    df['Timestamp'] = pd.to_datetime(df['Timestamp'], format='%d/%m/%y, %H.%M', errors='coerce')
    df.dropna(subset=['Timestamp'], inplace=True) 
    
    return df

# --- BAGIAN EKSEKUSI UTAMA ---

# !!! PENTING: Ganti path font ini dengan path font Arab yang ada di sistem Anda
# Contoh untuk Windows: 'C:/Windows/Fonts/arial.ttf'
# Contoh untuk MacOS: '/System/Library/Fonts/Supplemental/Arial.ttf'
# Contoh untuk Linux: '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'
# Jika tidak ditemukan, Word Cloud tidak akan dibuat.
font_path_arabic = 'C:/Windows/Fonts/arial.ttf' # Ganti dengan path yang benar jika perlu


folder_path = 'data-whatsapp'
df_full = parse_whatsapp_chat(folder_path)

if df_full.empty:
    print("Tidak ada data yang berhasil dibaca atau diproses.")
    print("Pastikan pola Regex di dalam kode sesuai dengan format chat di file .txt Anda.")
else:
    print("Data berhasil di-parsing! Jumlah pesan yang terdeteksi:", len(df_full))

    # --- FILTER DATA ---
    # Memfilter data untuk tidak menyertakan 'Nusa Restoria' agar hanya data pembeli yang dianalisis
    # df_full digunakan untuk analisis yang butuh konteks penjual dan pembeli
    df = df_full[df_full['Sender'] != 'Nusa Restoria']
    print(f"Data setelah memfilter 'Nusa Restoria' (hanya menampilkan data pembeli): {len(df)} pesan.")

    # --- ANALISIS DATA ---
    print("\nMemulai analisis data...")

    df['Hour'] = df['Timestamp'].dt.hour
    df['Day'] = df['Timestamp'].dt.day_name()
    df['Date'] = df['Timestamp'].dt.date

    sender_counts = df['Sender'].value_counts()
    top_10_senders = sender_counts.head(10)
    print("\n" + "="*30)
    print("     TOP 10 PENGIRIM PESAN TERBANYAK")
    print("="*30)
    print(top_10_senders)

    day_order = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    daily_activity = df['Day'].value_counts().reindex(day_order)
    print("\n" + "="*30)
    print("        AKTIVITAS PESAN PER HARI")
    print("="*30)
    print(daily_activity)

    date_activity = df['Date'].value_counts().sort_index()
    print("\n" + "="*30)
    print("      AKTIVITAS PESAN PER TANGGAL")
    print("="*30)
    print(date_activity)


    # --- EKSPOR DATA KE CSV ---
    print("\n\nMengekspor hasil analisis ke folder 'hasil_analisis_csv'...")
    output_folder = 'hasil_analisis_csv'
    os.makedirs(output_folder, exist_ok=True)

    # 1. Ekspor semua data chat yang sudah di-parse
    df.to_csv(os.path.join(output_folder, 'semua_pesan.csv'), index=False, encoding='utf-8-sig')
    print(f"- File 'semua_pesan.csv' berhasil disimpan.")

    # 2. Ekspor top 10 pengirim
    df_top_senders = top_10_senders.reset_index()
    df_top_senders.columns = ['Pengirim', 'Jumlah Pesan']
    df_top_senders.to_csv(os.path.join(output_folder, 'top_10_pengirim.csv'), index=False, encoding='utf-8-sig')
    print(f"- File 'top_10_pengirim.csv' berhasil disimpan.")

    # 3. Ekspor aktivitas harian
    df_daily_activity = daily_activity.reset_index()
    df_daily_activity.columns = ['Hari', 'Jumlah Pesan']
    df_daily_activity.to_csv(os.path.join(output_folder, 'aktivitas_harian.csv'), index=False, encoding='utf-8-sig')
    print(f"- File 'aktivitas_harian.csv' berhasil disimpan.")

    # 4. Ekspor aktivitas per tanggal
    df_date_activity = date_activity.reset_index()
    df_date_activity.columns = ['Tanggal', 'Jumlah Pesan']
    df_date_activity.to_csv(os.path.join(output_folder, 'aktivitas_per_tanggal.csv'), index=False, encoding='utf-8-sig')
    print(f"- File 'aktivitas_per_tanggal.csv' berhasil disimpan.")

    # --- VISUALISASI DATA ---
    print("\nMembuat visualisasi data...")
    sns.set_style("whitegrid")

    # Mengatur font yang mendukung karakter Arab untuk semua grafik.
    # Ini adalah perbaikan penting agar nama Arab tidak error atau tampil sebagai kotak-kotak.
    # Pastikan Anda memiliki font seperti 'Arial' atau 'Tahoma' yang terinstall.
    try:
        plt.rcParams['font.family'] = 'Arial'
        print("-> Font 'Arial' berhasil diatur untuk semua grafik.")
    except RuntimeError:
        print("-> Peringatan: Font 'Arial' tidak ditemukan. Label berbahasa Arab mungkin tidak akan tampil dengan benar.")
        print("   Silakan ganti dengan nama font lain yang terinstall di sistem Anda (misal: 'Tahoma').")

    # Grafik 1: Top 10 Pengirim
    plt.figure(figsize=(12, 8))
    
    # Reshape label nama pengirim agar teks Arab tampil benar
    reshaped_labels = [get_display(arabic_reshaper.reshape(label)) for label in top_10_senders.index]
    
    # Menggunakan bar chart horizontal agar nama yang panjang lebih mudah dibaca
    sns.barplot(x=top_10_senders.values, y=reshaped_labels, palette='viridis')
    plt.title('Top 10 Anggota Paling Aktif', fontsize=16)
    plt.xlabel('Jumlah Pesan', fontsize=12)
    plt.ylabel('Nama Pengirim', fontsize=12)
    plt.tight_layout()

    # Grafik 2: Aktivitas Chat per Jam
    plt.figure(figsize=(12, 6))
    hourly_activity = df['Hour'].value_counts().sort_index()
    hourly_activity.plot(kind='line', marker='o', color='coral')
    plt.title('Distribusi Pesan Sepanjang Hari', fontsize=16)
    plt.xlabel('Jam', fontsize=12)
    plt.ylabel('Jumlah Pesan', fontsize=12)
    plt.xticks(range(0, 24))
    plt.grid(True)
    plt.tight_layout()

    # Grafik 3: Aktivitas Pesan per Tanggal (dengan Nama Hari)
    plt.figure(figsize=(15, 7))
    ax = date_activity.plot(kind='line', marker='o', color='purple')
    plt.title('Aktivitas Pesan per Tanggal', fontsize=16)
    plt.xlabel('Tanggal', fontsize=12)
    plt.ylabel('Jumlah Pesan', fontsize=12)

    # Buat label kustom dengan nama hari dalam Bahasa Indonesia
    day_map_en_to_id = {
        'Monday': 'Senin', 'Tuesday': 'Selasa', 'Wednesday': 'Rabu',
        'Thursday': 'Kamis', 'Friday': 'Jumat', 'Saturday': 'Sabtu', 'Sunday': 'Minggu'
    }
    
    # Konversi index (objek datetime.date) ke pandas Timestamps untuk menggunakan .dt
    timestamps = pd.to_datetime(date_activity.index)
    
    # Buat label baru dalam format "Hari, DD-MM-YY"
    new_labels = [f"{day_map_en_to_id[ts.day_name()]}, {ts.strftime('%d-%m-%y')}" for ts in timestamps]
    
    # Atur label kustom pada sumbu-x
    ax.set_xticks(timestamps)
    ax.set_xticklabels(new_labels, rotation=45, ha='right')
    
    plt.grid(True)
    plt.tight_layout()

    # --- Grafik 5: Analisis Jenis Gaharu yang Dicari ---
    print("\nMembuat grafik analisis jenis gaharu yang dicari...")
    
    gaharu_keywords = {
        'Scented Bakhoor (بخور معطر)': ['بخور معطر'],
        'Natural (طبيعي)': ['طبيعي', 'الطبيعي'],
        'Enhanced (محسن)': ['محسن', 'المحسن'],
        'Artificial (صناعي)': ['صناعي', 'الصناعي'],
        'Sumatra (سومطرة)': ['سومطرة', 'sumatra'],
        'Kalimantan (كاليمانتان)': ['كاليمانتان', 'kalimantan'],
        'Merauke (ميروكي)': ['ميروكي', 'merauke']
    }
    
    gaharu_counts = {key: 0 for key in gaharu_keywords.keys()}
    
    # Menggunakan pesan dari pembeli saja (df sudah difilter)
    for message in df['Message']:
        # Pastikan pesan adalah string untuk menghindari error
        if isinstance(message, str):
            for gaharu_type, keywords in gaharu_keywords.items():
                if any(keyword in message for keyword in keywords):
                    gaharu_counts[gaharu_type] += 1
    
    # Filter jenis gaharu yang tidak pernah disebutkan
    gaharu_counts = {k: v for k, v in gaharu_counts.items() if v > 0}
    
    if gaharu_counts:
        # Membuat Series Pandas untuk di-plot
        gaharu_series = pd.Series(gaharu_counts).sort_values(ascending=True)
        
        plt.figure(figsize=(12, 8))
        
        # Reshape label agar teks Arab tampil benar
        reshaped_labels = [get_display(arabic_reshaper.reshape(label)) for label in gaharu_series.index]
        
        sns.barplot(x=gaharu_series.values, y=reshaped_labels, palette='plasma')
        plt.title('Jenis Gaharu yang Paling Sering Disebutkan', fontsize=16)
        plt.xlabel('Jumlah Penyebutan', fontsize=12)
        plt.ylabel('Jenis Gaharu', fontsize=12)
        # Menambahkan label angka di ujung setiap bar
        for index, value in enumerate(gaharu_series.values):
            plt.text(value, index, f' {value}', va='center', ha='left')
        plt.tight_layout()
    else:
        print("-> Tidak ditemukan penyebutan jenis gaharu spesifik dalam chat.")

    # --- Grafik 6: Analisis Pertanyaan Umum ---
    print("\nMembuat grafik analisis pertanyaan umum dari pelanggan...")

    # Kata kunci untuk pertanyaan umum (dalam bentuk lowercase untuk pencocokan)
    question_keywords = {
        'Tanya Harga (كم سعر)': ['كم سعر', 'كم السعر', 'الأسعار', 'كم سعره', 'كم عر'],
        'Tanya Detail Produk (تفاصيل)': ['تفاصيل', 'ايش العروض', 'ماهي الاعواد', 'اي نوع', 'ما هو النوع'],
        'Tanya Ketersediaan (هل متوفر)': ['هل يوجد', 'هل عندكم', 'متوفر'],
        'Minta Gambar/Katalog (صور)': ['صور', 'العرض', 'كاتلوج'],
        'Tanya Lokasi/Pengiriman (المكان فين)': ['المكان فين', 'كيف نوصل', 'كيف ترسل']
    }

    question_counts = {key: 0 for key in question_keywords.keys()}

    # Iterasi hanya pada pesan dari pelanggan (df)
    for message in df['Message']:
        if isinstance(message, str):
            msg_lower = message.lower()
            # Iterasi pada setiap kategori pertanyaan
            for category, keywords in question_keywords.items():
                # Jika ada kata kunci yang cocok, hitung dan hentikan pencarian untuk pesan ini
                if any(keyword in msg_lower for keyword in keywords):
                    question_counts[category] += 1
                    break # Agar satu pesan tidak dihitung di beberapa kategori

    # Filter kategori yang tidak pernah ditanyakan
    question_counts = {k: v for k, v in question_counts.items() if v > 0}

    if question_counts:
        question_series = pd.Series(question_counts).sort_values(ascending=True)
        plt.figure(figsize=(12, 8))
        reshaped_labels = [get_display(arabic_reshaper.reshape(label)) for label in question_series.index]
        sns.barplot(x=question_series.values, y=reshaped_labels, palette='crest')
        plt.title('Top 5 Kategori Pertanyaan Umum dari Pelanggan', fontsize=16)
        plt.xlabel('Jumlah Pertanyaan', fontsize=12)
        plt.ylabel('Kategori Pertanyaan', fontsize=12)
        for index, value in enumerate(question_series.values):
            plt.text(value, index, f' {value}', va='center', ha='left')
        plt.tight_layout()
    else:
        print("-> Tidak ditemukan pertanyaan umum yang cocok dengan kata kunci.")

    # --- Grafik 7: Analisis Penyebab Chat Tidak Dibalas ---
    print("\nMembuat grafik analisis penyebab chat tidak dibalas...")

    unreplied_chats = []
    # Group by conversation file, menggunakan df_full yang berisi semua chat
    for filename, group in df_full.groupby('Filename'):
        group = group.sort_values('Timestamp')
        if not group.empty:
            # Cek apakah pesan terakhir dalam sebuah percakapan dikirim oleh penjual
            if group.iloc[-1]['Sender'] == 'Nusa Restoria':
                unreplied_chats.append(group)

    if unreplied_chats:
        drop_off_reasons = Counter()
        
        # Definisikan kategori dan kata kunci (urutan penting, dari spesifik ke umum)
        categories = {
            'Follow-up (Tidak Dibalas)': ['متابعة بسيطة'],
            'Diskusi Harga': ['سعر', 'أسعار', 'price', 'كم جراما', 'ريال'],
            'Ajakan Meeting/Call': ['اتصال', 'gmeet', 'meet.google.com', 'اجتماع', 'رابط'],
            'Permintaan Alamat/Info Pengiriman': ['الاسم الكامل', 'الدولة والمدينة', 'عنوان', 'شحن', 'الرمز البريدي', 'بيانات'],
            'Diskusi Sampel': ['عينة', 'sample'],
            'Diskusi Detail Produk': ['محسن', 'طبيعي', 'صناعي', 'سومطرة', 'كاليمانتان', 'ميروكي', 'نوع', 'تفاصيل']
        }

        for chat in unreplied_chats:
            # Konteks adalah 5 pesan terakhir dalam percakapan
            context_messages = " ".join(chat.tail(5)['Message']).lower()
            
            categorized = False
            for category, keywords in categories.items():
                if any(kw in context_messages for kw in keywords):
                    drop_off_reasons[category] += 1
                    categorized = True
                    break # Pindah ke chat berikutnya setelah dikategorikan
            
            if not categorized:
                drop_off_reasons['Lain-lain / Minat Awal Rendah'] += 1
        
        # Membuat visualisasi
        if drop_off_reasons:
            reasons_series = pd.Series(drop_off_reasons).sort_values(ascending=True)
            
            plt.figure(figsize=(12, 8))
            
            reshaped_labels = [get_display(arabic_reshaper.reshape(label)) for label in reasons_series.index]
            
            sns.barplot(x=reasons_series.values, y=reshaped_labels, palette='mako')
            plt.title('Analisis Potensi Penyebab Pembeli Tidak Membalas', fontsize=16)
            plt.xlabel('Jumlah Chat', fontsize=12)
            plt.ylabel('Kategori Konteks Terakhir', fontsize=12)
            for index, value in enumerate(reasons_series.values):
                plt.text(value, index, f' {value}', va='center', ha='left')
            plt.tight_layout()
    else:
        print("-> Tidak ditemukan chat yang tidak dibalas oleh pembeli.")

    # --- Grafik 8: Analisis Waktu Respons Penjual ---
    print("\nMembuat grafik analisis waktu respons...")

    response_times_minutes = []
    # Group by conversation file, menggunakan df_full yang berisi semua chat
    for filename, group in df_full.groupby('Filename'):
        group = group.sort_values('Timestamp')
        
        # Cari pesan dari pelanggan yang diikuti oleh balasan dari 'Nusa Restoria'
        for i in range(len(group) - 1):
            current_sender = group.iloc[i]['Sender']
            next_sender = group.iloc[i+1]['Sender']
            
            # Cek jika pengirim saat ini BUKAN 'Nusa Restoria' dan pengirim berikutnya ADALAH 'Nusa Restoria'
            if current_sender != 'Nusa Restoria' and next_sender == 'Nusa Restoria':
                time_diff = group.iloc[i+1]['Timestamp'] - group.iloc[i]['Timestamp']
                if time_diff.total_seconds() > 0:
                    response_times_minutes.append(time_diff.total_seconds() / 60) # dalam menit
                    # Hanya ambil respons pertama yang valid per pelanggan dalam satu sesi
                    break 

    if response_times_minutes:
        # Hapus outlier untuk visualisasi yang lebih baik (misal, respons lebih dari 6 jam)
        response_times_filtered = [t for t in response_times_minutes if t < 360] 
        
        plt.figure(figsize=(12, 7))
        sns.histplot(response_times_filtered, bins=50, kde=True, color='skyblue')
        plt.title('Distribusi Waktu Respons Penjual (dalam Menit)', fontsize=16)
        plt.xlabel('Waktu Respons (Menit)', fontsize=12)
        plt.ylabel('Jumlah Kejadian', fontsize=12)
        avg_response_time = pd.Series(response_times_filtered).mean()
        plt.axvline(avg_response_time, color='red', linestyle='--', label=f'Rata-rata: {avg_response_time:.2f} menit')
        plt.legend()
        plt.tight_layout()
    else:
        print("-> Tidak cukup data untuk menganalisis waktu respons.")

    # --- Grafik 9: Corong Konversi Pelanggan ---
    print("\nMembuat grafik corong konversi pelanggan...")

    # Tahapan funnel dan kata kunci (dalam huruf kecil)
    funnel_stages = {
        '1. Kontak Awal': {'count': 0},
        '2. Diskusi Lanjut (Harga/Jenis)': {'count': 0, 'keywords': ['سعر', 'كم', 'نوع', 'تفاصيل', 'محسن', 'طبيعي', 'صناعي', 'price', 'عينة', 'sample', 'شحن', 'ارسال', 'أسعار']},
        '3. Potensi Konversi (Kirim Alamat)': {'count': 0, 'keywords': ['لتجهيز طلبك', 'الاسم الكامل', 'الدولة والمدينة', 'عنوان', 'بيانات']}
    }

    conversations = df_full.groupby('Filename')

    # Stage 1: Semua percakapan unik dihitung
    funnel_stages['1. Kontak Awal']['count'] = len(conversations)

    for filename, group in conversations:
        seller_messages_text = " ".join(group[group['Sender'] == 'Nusa Restoria']['Message'].astype(str)).lower()
        full_chat_text = " ".join(group['Message'].astype(str)).lower()

        # Stage 2: Ada diskusi lanjut
        if any(kw in full_chat_text for kw in funnel_stages['2. Diskusi Lanjut (Harga/Jenis)']['keywords']):
            funnel_stages['2. Diskusi Lanjut (Harga/Jenis)']['count'] += 1
        
        # Stage 3: Penjual mengirimkan form pemesanan
        if any(kw in seller_messages_text for kw in funnel_stages['3. Potensi Konversi (Kirim Alamat)']['keywords']):
            funnel_stages['3. Potensi Konversi (Kirim Alamat)']['count'] += 1

    # Membuat Series untuk plot
    funnel_counts = {stage: data['count'] for stage, data in funnel_stages.items()}
    funnel_series = pd.Series(funnel_counts)

    if not funnel_series.empty:
        plt.figure(figsize=(12, 8))
        ax = sns.barplot(x=funnel_series.values, y=funnel_series.index, palette='magma')
        plt.title('Corong Konversi Pelanggan (Customer Funnel)', fontsize=16)
        plt.xlabel('Jumlah Percakapan', fontsize=12)
        plt.ylabel('Tahapan Funnel', fontsize=12)

        # Tambahkan label persentase dan angka absolut
        total_conversations = funnel_series.iloc[0]
        for i, count in enumerate(funnel_series.values):
            if total_conversations > 0:
                percentage_of_total = (count / total_conversations) * 100
                ax.text(count, i, f' {count} ({percentage_of_total:.1f}%)', va='center', ha='left', fontsize=12)
        
        plt.tight_layout()
    else:
        print("-> Tidak ada data percakapan untuk membuat funnel.")

    # --- Grafik 10: Distribusi Geografis Pelanggan ---
    print("\nMembuat grafik distribusi geografis pelanggan...")

    # Daftar kota utama di Saudi dan negara lain yang mungkin disebutkan
    locations = {
        'Riyadh (رياض)': ['رياض'],
        'Jeddah (جدة)': ['جدة', 'جده'],
        'Makkah (مكة)': ['مكة'],
        'Tabuk (تبوك)': ['تبوك'],
        'Kuwait (الكويت)': ['الكويت', 'kuwait'],
    }
    location_counts = {loc: 0 for loc in locations.keys()}

    # Cari penyebutan kota dalam pesan pelanggan (df)
    for message in df['Message']:
        if isinstance(message, str):
            msg_lower = message.lower()
            for location, keywords in locations.items():
                if any(keyword in msg_lower for keyword in keywords):
                    location_counts[location] += 1
                    break 

    location_counts = {k: v for k, v in location_counts.items() if v > 0}

    if location_counts:
        location_series = pd.Series(location_counts).sort_values(ascending=True)
        
        plt.figure(figsize=(12, 7))
        reshaped_labels = [get_display(arabic_reshaper.reshape(label)) for label in location_series.index]
        sns.barplot(x=location_series.values, y=reshaped_labels, palette='cubehelix')
        plt.title('Distribusi Geografis Pelanggan Berdasarkan Penyebutan Lokasi', fontsize=16)
        plt.xlabel('Jumlah Penyebutan', fontsize=12)
        plt.ylabel('Lokasi', fontsize=12)
        for index, value in enumerate(location_series.values):
            plt.text(value, index, f' {value}', va='center', ha='left')
        plt.tight_layout()
    else:
        print("-> Tidak ditemukan penyebutan lokasi spesifik oleh pelanggan.")

    # --- Grafik 11: Distribusi Jumlah Pesan per Percakapan ---
    print("\nMembuat grafik distribusi jumlah pesan per percakapan...")

    message_counts_per_convo = df_full.groupby('Filename').size()

    if not message_counts_per_convo.empty:
        plt.figure(figsize=(12, 7))
        # Filter untuk visualisasi yang lebih baik, misal percakapan < 50 pesan
        sns.histplot(message_counts_per_convo[message_counts_per_convo < 50], bins=25, kde=False, color='teal')
        plt.title('Distribusi Jumlah Pesan per Percakapan', fontsize=16)
        plt.xlabel('Jumlah Pesan dalam Satu Percakapan', fontsize=12)
        plt.ylabel('Jumlah Percakapan', fontsize=12)
        plt.tight_layout()
    else:
        print("-> Tidak ada data untuk menganalisis jumlah pesan per percakapan.")


    # Menampilkan semua grafik
    print("\nMenampilkan grafik... Tutup semua jendela grafik untuk mengakhiri program.")
    plt.show()