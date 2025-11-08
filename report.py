# report.py

import pandas as pd
from jinja2 import Environment, FileSystemLoader
import datetime

# --- 1. Veri Yükleme ---
def load_data():
    try:
        # Excel dosyalarını oku
        df_envanter = pd.read_excel('data_input/envanter.xlsx')
        df_ariza = pd.read_excel('data_input/ariza_kayitlari.xlsx')
        return df_envanter, df_ariza
    except FileNotFoundError:
        print("HATA: Gerekli Excel dosyaları 'data_input' klasöründe bulunamadı.")
        return None, None

# --- 2. Veri İşleme ve Analiz ---
def process_data(df_envanter, df_ariza):
    # Temel istatistikler
    stats = {
        'toplam_cihaz': len(df_envanter),
        'cihaz_durum_dagilimi': df_envanter['Durum'].value_counts().to_dict(),
        'cihaz_tur_dagilimi': df_envanter['CihazTuru'].value_counts().to_dict(),
        'toplam_ariza': len(df_ariza),
        'acik_ariza_sayisi': len(df_ariza[df_ariza['Durum'] == 'Açık'])
    }

    # "Açık" durumdaki arızaları ve envanter bilgilerini birleştir
    acik_arizalar_detayli = pd.merge(
        df_ariza[df_ariza['Durum'] == 'Açık'],
        df_envanter,
        on='CihazID',
        how='left' # Sadece arıza kaydı olanları al
    )
    
    # Jinja'da kolay kullanmak için list of dicts formatına çevir
    context = {
        'stats': stats,
        'acik_arizalar_listesi': acik_arizalar_detayli.to_dict('records'),
        'tum_envanter_listesi': df_envanter.to_dict('records'),
        'rapor_tarihi': datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    
    return context

# report.py (devamı)

# --- 3. Raporu Oluşturma (Jinja2) ---
def create_report(context):
    # Jinja ortamını ayarla (template'lerin olduğu klasörü belirt)
    env = Environment(loader=FileSystemLoader('templates'))
    template = env.get_template('template.html')
    
    # Veriyi template ile birleştir
    html_output = template.render(context)
    
    # Oluşturulan HTML'i dosyaya yaz
    output_filename = f"output/bit_raporu_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.html"
    with open(output_filename, 'w', encoding='utf-8') as f:
        f.write(html_output)
        
    print(f"Rapor başarıyla oluşturuldu: {output_filename}")


# --- 4. Ana Çalıştırma Bloğu ---
def main():
    print("Rapor oluşturuluyor...")
    df_envanter, df_ariza = load_data()
    
    if df_envanter is not None:
        context = process_data(df_envanter, df_ariza)
        create_report(context)

if __name__ == "__main__":
    main()