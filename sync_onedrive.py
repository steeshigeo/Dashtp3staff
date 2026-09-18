import os
import re
import sys
import json
import requests
import pandas as pd

# Link OneDrive
URL_MERCHANDISE_REPORT = "https://1drv.ms/x/c/426713E72C49EDEB/IQB6fEpkF1TaRr1Knh3LPIqaAWi45hl7dda2wv6TXqeRxi0?e=LZvUgj"
URL_SALESPERSON_WISE = "https://1drv.ms/x/c/426713E72C49EDEB/IQBD3oB0QQCkTZRPXh7HhhWoAQ8oQIHav9BkCWSKDtwxFXI?e=YjQH93"

def get_download_url(share_url):
    """Memastikan parameter download=1 ada untuk mengunduh langsung dari OneDrive."""
    if not share_url:
        return ""
    if "download=1" not in share_url:
        if "?" in share_url:
            return f"{share_url}&download=1"
        else:
            return f"{share_url}?download=1"
    return share_url

def download_file(url, output_path):
    download_url = get_download_url(url)
    print(f"[*] Mengunduh: {output_path} dari {download_url}")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8'
    }
    try:
        session = requests.Session()
        response = session.get(download_url, headers=headers, allow_redirects=True, timeout=30)
        if response.status_code == 200:
            with open(output_path, 'wb') as f:
                f.write(response.content)
            print(f"[✓] Berhasil mengunduh {output_path} ({len(response.content)} bytes)")
            return True
        else:
            print(f"[!] Gagal mengunduh {output_path}. Status code: {response.status_code}")
            return False
    except Exception as e:
        print(f"[!] Exception saat mengunduh {output_path}: {e}")
        return False

def build_merchandise_catalog(merch_file):
    """
    Parses merchandisesalesreport Excel.
    Kolom C (index 2): Product Group ('DEVICES', 'ACCESSORIES', dll)
    Kolom F (index 5): Description / Nama Produk
    Juga mencatat kode Artikel (Kolom B/D) jika ada.
    """
    catalog = {}
    if not os.path.exists(merch_file):
        print(f"[!] File catalog {merch_file} tidak ditemukan, menggunakan catalog fallback.")
        return catalog

    try:
        df = pd.read_excel(merch_file, header=None)
        print(f"[*] Parsing Merchandise Report: {df.shape}")
        
        for idx, row in df.iterrows():
            col_c = str(row[2]).strip().upper() if len(row) > 2 and pd.notna(row[2]) else ""
            col_f = str(row[5]).strip().upper() if len(row) > 5 and pd.notna(row[5]) else ""
            col_b = str(row[1]).strip().upper() if len(row) > 1 and pd.notna(row[1]) else ""
            col_d = str(row[3]).strip().upper() if len(row) > 3 and pd.notna(row[3]) else ""

            cat = None
            if "DEVICES" in col_c:
                cat = "Device"
            elif "ACCESSORIES" in col_c:
                cat = "Accessories"

            if cat:
                if col_f:
                    catalog[col_f] = cat
                if col_b:
                    catalog[col_b] = cat
                if col_d:
                    catalog[col_d] = cat

        print(f"[✓] Catalog berhasil dibuat dengan {len(catalog)} mapping produk.")
    except Exception as e:
        print(f"[!] Error parsing Merchandise Report: {e}")
    
    return catalog

def parse_salespersonwise(sp_file, catalog):
    print(f"[*] Parsing Salespersonwise File: {sp_file}")
    df = pd.read_excel(sp_file, header=None)
    
    records = []
    current_date = None
    current_staff = None
    
    vas_keywords = ['klabronze', 'klasilver', 'klagold', 'klaemerald', 'kladiamond', 'klaplatinum', 'klatitanium', 'klavvip', 'kla', 'tsl', 'idt', 'xxl']
    
    acc_exception_words = [
        'keyboard', 'sleeve', 'case', 'adapter', 'cable', 'guard', 'protector', 
        'cover', 'strap', 'band', 'tempered', 'glass', 'film', 'battery', 
        'charger', 'stand', 'pouch', 'holder', 'clearvue', 'combat', 'vero', 'm-neo', 'card', 'bag'
    ]
    
    device_keywords = ['iphone', 'ipad', 'macbook', 'imac', 'macmini', 'mac mini', 'macpro', 'mac pro', 'mba', 'mbp', 'mbn', 'apple watch', 'watch ultra', 'watch s', 'watch 11', 'watch 10', 'watch se']
    
    i = 0
    n = len(df)
    
    while i < n:
        c0 = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
        c1 = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ''
        
        # Deteksi Tanggal (DD-MM-YYYY)
        if re.match(r'^\d{2}-\d{2}-\d{4}$', c0):
            current_date = c0
            i += 1
            continue
            
        # Deteksi Staff
        if '/' in c0 and not c0.startswith('Total For') and not re.match(r'^\d{2}-', c0):
            parts = c0.split('/')
            if len(parts) == 2 and parts[0].strip().isdigit():
                current_staff = c0.strip()
                i += 1
                continue
                
        # Deteksi Total Staff
        if c0.startswith('Total For'):
            current_staff = None
            i += 1
            continue
            
        # Deteksi Produk Transaksi
        if c0 and current_staff and current_date:
            product_name = c0
            article = c1
            i += 1
            
            while i < n:
                rc0 = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
                rc1 = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ''
                rc2 = str(df.iloc[i, 2]).strip() if pd.notna(df.iloc[i, 2]) else ''
                
                if rc0.startswith('Total For') or ('/' in rc0 and rc0.split('/')[0].strip().isdigit()) or re.match(r'^\d{2}-\d{2}-\d{4}$', rc0):
                    break
                    
                try:
                    price = float(rc0)
                    net_price = float(rc2) if rc2 and rc2.replace('.', '', 1).isdigit() else price
                    
                    if net_price > 0:
                        p_upper = product_name.upper()
                        art_upper = article.upper()
                        col_c_upper = rc2.upper()
                        
                        # 1. CEK VAS (Prioritas Utama)
                        category = "Accessories"
                        if any(kw in art_upper or kw in p_upper for kw in [k.upper() for k in vas_keywords]):
                            category = "VAS"
                        # 2. MATCH DENGAN CATALOG MERCHANDISE REPORT (KOLOM C & F)
                        elif p_upper in catalog:
                            category = catalog[p_upper]
                        elif art_upper in catalog:
                            category = catalog[art_upper]
                        # 3. FALLBACK: PRODUCT GROUP KOLOM C DI SALESPERSONWISE
                        elif 'DEVICES' in col_c_upper:
                            category = "Device"
                        elif 'ACCESSORIES' in col_c_upper:
                            category = "Accessories"
                        # 4. FALLBACK: KATA KUNCI AKSESORI & DEVICE
                        elif any(re.search(r'\b' + re.escape(aw.upper()) + r'\b', p_upper) for aw in acc_exception_words):
                            category = "Accessories"
                        elif any(dk.upper() in p_upper for dk in device_keywords):
                            category = "Device"
                            
                        # MAPPING LOB FOCUS
                        lob_focus = None
                        if 'IPHONE 15' in p_upper:
                            lob_focus = 'iPhone 15'
                        elif 'IPAD' in p_upper and ('11TH' in p_upper or '10TH' in p_upper or 'A16' in p_upper):
                            lob_focus = 'iPad 11th'
                        elif 'MBN' in p_upper or 'MACBOOK' in p_upper or 'MBA' in p_upper or 'MBP' in p_upper:
                            lob_focus = 'MBN'
                        elif 'APPLE WATCH' in p_upper or 'WATCH' in p_upper or 'AW' in p_upper:
                            lob_focus = 'AW'
                        elif 'AIRPOD' in p_upper:
                            lob_focus = 'AirPods'

                        records.append({
                            'date': current_date,
                            'staff': current_staff,
                            'product_name': product_name,
                            'article': article,
                            'price': price,
                            'receipt': rc1,
                            'net_price': net_price,
                            'category': category,
                            'lob_focus': lob_focus
                        })
                    i += 1
                except ValueError:
                    break
            continue
            
        i += 1
        
    print(f"[✓] Berhasil memproses {len(records)} transaksi sales.")
    return records

def main():
    url_merch = os.environ.get("URL_MERCHANDISE_REPORT", URL_MERCHANDISE_REPORT)
    url_sales = os.environ.get("URL_SALESPERSON_WISE", URL_SALESPERSON_WISE)
    
    file_merch = "merchandisesalesreport.xlsx"
    file_sales = "salespersonwise.xls"
    json_filename = "data.json"
    
    # Unduh file
    download_file(url_merch, file_merch)
    if download_file(url_sales, file_sales):
        catalog = build_merchandise_catalog(file_merch)
        sales_data = parse_salespersonwise(file_sales, catalog)
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(sales_data, f, indent=2, ensure_ascii=False)
        print(f"[✓] Data JSON diperbarui ke {json_filename}")
    else:
        print("[!] Gagal mengunduh file sales, menghentikan eksekusi.")
        sys.exit(1)

if __name__ == "__main__":
    main()
