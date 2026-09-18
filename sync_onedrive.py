import os
import re
import sys
import json
import base64
import requests
import pandas as pd

# Link OneDrive default milik Anda
DEFAULT_ONEDRIVE_URL = "https://1drv.ms/x/c/426713E72C49EDEB/IQBD3oB0QQCkTZRPXh7HhhWoAQ8oQIHav9BkCWSKDtwxFXI?e=zoQ86E"

def get_direct_download_url(share_url):
    """Mengonversi link sharing OneDrive menjadi direct download link."""
    try:
        encoded = base64.b64encode(share_url.encode('utf-8')).decode('utf-8')
        encoded_safe = encoded.rstrip('=').replace('/', '_').replace('+', '-')
        return f"https://api.onedrive.com/v1.0/shares/u!{encoded_safe}/root/content"
    except Exception as e:
        print(f"[!] Error converting URL: {e}")
        return share_url

def download_excel(url, output_path="salespersonwise.xls"):
    direct_url = get_direct_download_url(url)
    print(f"[*] Mengunduh data dari: {direct_url}")
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
    }
    response = requests.get(direct_url, headers=headers, allow_redirects=True)
    if response.status_code == 200:
        with open(output_path, 'wb') as f:
            f.write(response.content)
        print(f"[✓] Berhasil mengunduh {output_path} ({len(response.content)} bytes)")
        return True
    else:
        print(f"[!] Gagal mengunduh file. Status code: {response.status_code}")
        return False

def parse_sales_xls(file_path):
    print(f"[*] Parsing file Excel: {file_path}")
    df = pd.read_excel(file_path, header=None)
    
    records = []
    current_date = None
    current_staff = None
    
    vas_keywords = ['kla', 'tsl', 'idt', 'xxl']
    device_keywords = ['iphone', 'ipad', 'macbook', 'imac', 'mac mini', 'apple watch', 'watch ultra', 'watch s', 'watch 11', 'watch 10']
    
    i = 0
    n = len(df)
    
    while i < n:
        c0 = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
        c1 = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ''
        
        # Deteksi Baris Tanggal (DD-MM-YYYY)
        if re.match(r'^\d{2}-\d{2}-\d{4}$', c0):
            current_date = c0
            i += 1
            continue
            
        # Deteksi Baris Staff
        if '/' in c0 and not c0.startswith('Total For') and not re.match(r'^\d{2}-', c0):
            parts = c0.split('/')
            if len(parts) == 2 and parts[0].strip().isdigit():
                current_staff = c0.strip()
                i += 1
                continue
                
        # Deteksi Baris Total Staff
        if c0.startswith('Total For'):
            current_staff = None
            i += 1
            continue
            
        # Deteksi Blok Produk
        if c0 and current_staff and current_date:
            product_name = c0
            article = c1
            i += 1
            
            while i < n:
                rc0 = str(df.iloc[i, 0]).strip() if pd.notna(df.iloc[i, 0]) else ''
                rc1 = str(df.iloc[i, 1]).strip() if pd.notna(df.iloc[i, 1]) else ''
                rc2 = df.iloc[i, 2] if pd.notna(df.iloc[i, 2]) else None
                
                if rc0.startswith('Total For') or ('/' in rc0 and rc0.split('/')[0].strip().isdigit()) or re.match(r'^\d{2}-\d{2}-\d{4}$', rc0):
                    break
                    
                try:
                    price = float(rc0)
                    net_price = float(rc2) if rc2 is not None else price
                    
                    p_lower = product_name.lower()
                    art_lower = article.lower()
                    
                    category = "Accessories"
                    if any(kw in art_lower or kw in p_lower for kw in vas_keywords):
                        category = "VAS"
                    elif any(dk in p_lower for dk in device_keywords):
                        category = "Device"
                        
                    lob_focus = None
                    if 'iphone 15' in p_lower:
                        lob_focus = 'iPhone 15'
                    elif 'ipad' in p_lower and ('a16' in p_lower or '10th' in p_lower or '10.9' in p_lower):
                        lob_focus = 'iPad A16'
                    elif 'mbn' in p_lower or 'macbook' in p_lower:
                        lob_focus = 'MBN'
                    elif 'aw se' in p_lower or 'watch se' in p_lower:
                        lob_focus = 'AW SE'
                    elif 'airpods' in p_lower:
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
    # Menggunakan DEFAULT_ONEDRIVE_URL jika variabel environment kosong
    onedrive_url = os.environ.get("ONEDRIVE_URL")
    if not onedrive_url or onedrive_url.strip() == "":
        onedrive_url = DEFAULT_ONEDRIVE_URL
        
    xls_filename = "salespersonwise.xls"
    json_filename = "data.json"
    
    if download_excel(onedrive_url, xls_filename):
        sales_data = parse_sales_xls(xls_filename)
        with open(json_filename, 'w', encoding='utf-8') as f:
            json.dump(sales_data, f, indent=2, ensure_ascii=False)
        print(f"[✓] Data JSON diperbarui ke {json_filename}")
    else:
        print("[!] Download gagal, menghentikan eksekusi.")
        sys.exit(1)

if __name__ == "__main__":
    main()
