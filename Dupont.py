import pandas as pd

def calculate_dupont_analysis(input_file, output_file):
    # Excel dosyasını sheet'lere göre yükle
    xls = pd.ExcelFile(input_file)
    
    # Sonuçları saklamak için boş bir sözlük
    dupont_results = {}
    
    # Her bir sheet (hisse) için işlemleri tekrarla
    for sheet_name in xls.sheet_names:
        try:
            # Sheet'i oku ve 'itemDescTr' sütununu index olarak ayarla
            df = pd.read_excel(input_file, sheet_name=sheet_name)
            df.set_index('itemDescTr', inplace=True)
        except Exception as e:
            print(f"Sheet '{sheet_name}' okunurken hata: {e}")
            continue

        # Çeyrek tarihleri içeren sütunların belirlenmesi
        # Varsayım: ilk iki sütun (örn. 'itemCode' ve 'itemDescEng') hariç kalan sütunlar çeyrek tarihlerini içeriyor
        quarter_columns = df.columns.tolist()[2:]
        
        # Gerekli finansal kalemleri al (satır isimleri)
        try:
            net_profit = df.loc["Dönem Net Kar/Zararı", quarter_columns]
            sales = df.loc["Satış Gelirleri", quarter_columns]
            total_assets = df.loc["TOPLAM VARLIKLAR", quarter_columns]
            equity = df.loc["Özkaynaklar", quarter_columns]
        except KeyError as e:
            print(f"'{e.args[0]}' kalemi sheet '{sheet_name}' içerisinde bulunamadı. İşlem atlanıyor.")
            continue

        # Verileri sayısal formata çevir (hatalı girişler NaN olarak işlenecek)
        net_profit = pd.to_numeric(net_profit, errors='coerce')
        sales = pd.to_numeric(sales, errors='coerce')
        total_assets = pd.to_numeric(total_assets, errors='coerce')
        equity = pd.to_numeric(equity, errors='coerce')
        
        # DuPont bileşenlerini hesapla
        profit_margin = net_profit / sales
        asset_turnover = sales / total_assets
        equity_multiplier = total_assets / equity
        
        # ROE hesaplaması (sonuç yüzde cinsinden olacak şekilde 100 ile çarpılıyor)
        roe = (profit_margin * asset_turnover * equity_multiplier) * 100
        
        # Sonuçları sözlüğe ekle (sheet adı hisse adı olarak kullanılacak)
        dupont_results[sheet_name] = roe

    # Sözlükten DataFrame oluştur; satırlarda hisse adları, sütunlarda çeyrek tarihleri yer alır
    result_df = pd.DataFrame(dupont_results).T

    # Sütunları kronolojik sıraya göre yeniden düzenle
    try:
        sorted_columns = sorted(result_df.columns, key=lambda x: pd.to_datetime(x, format='%Y/%m'))
        result_df = result_df.reindex(sorted_columns, axis=1)
    except Exception as e:
        print("Tarih sıralamasında hata oluştu, sütunlar olduğu gibi bırakılıyor:", e)
    
    # İlk hücrede 'Hisse Adı' yazan, sonrasında çeyrek tarihlerin olduğu bir tablo oluşturmak için index düzenleniyor
    result_df.index.name = "Hisse Adı"
    result_df.reset_index(inplace=True)
    
    # Sonuçları Excel dosyasına kaydet
    result_df.to_excel(output_file, index=False)
    print(f"DuPont analizi sonuçları '{output_file}' dosyasına kaydedildi.")

if __name__ == "__main__":
    input_excel = "finansallar.xlsx"  # Localde bulunan excel dosyası
    output_excel = "dupont_analysis_output.xlsx"
    calculate_dupont_analysis(input_excel, output_excel)
