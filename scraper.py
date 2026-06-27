import os
import json
import time
import requests
import pandas as pd

def get_nova_price(code):
    """Nova.ge-ს API-დან ფასის პირდაპირი წამოღება ბრაუზერის გარეშე"""
    try:
        # Nova-ს შიდა საძიებო მოთხოვნა
        url = f"https://nova.ge/index.php?dispatch=products.search&q={code}&subcats=Y&pcode_from_q=Y&pshort=Y&pfull=Y&pname=Y&pkeywords=Y&search_performed=Y"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
            "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8"
        }
        
        response = requests.get(url, headers=headers, timeout=15)
        
        # თუ პროდუქტი იპოვა, ტექსტში ვეძებთ ფასის კლასს
        if "ty-price-num" in response.text:
            part = response.text.split('class="ty-price-num">')[1]
            price_str = part.split('</span>')[0].replace("₾", "").replace(",", "").strip()
            return float(price_str)
    except Exception:
        pass
    return None

def get_liloshop_price(code):
    """Liloshop.ge-ს ძებნის API-დან ფასის წამოღება"""
    try:
        url = f"https://liloshop.ge/search?q={code}"
        headers = {
            "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
        }
        response = requests.get(url, headers=headers, timeout=15)
        
        if 'class="price"' in response.text:
            part = response.text.split('class="price">')[1]
            price_str = part.split('</span>')[0].replace("₾", "").replace(" ", "").replace(",", "").strip()
            return float(price_str)
    except Exception:
        pass
    return None

def main():
    input_file = "ai.xlsx"
    output_file = "fasebi_ganaxlebuli.xlsx"
    
    if not os.path.exists(input_file):
        print(f"⚠️ ფაილი {input_file} ვერ მოიძებნა.")
        return

    df = pd.read_excel(input_file)
    print(f"🚀 სკრაპინგი დაიწყო API რეჟიმში. დასამუშავებელია {len(df)} კოდი...")
    
    for index, row in df.iterrows():
        code = str(row['code'])
        print(f"🔄 [{index + 1}/{len(df)}] კოდი: {code}...", end="", flush=True)
        
        # Nova
        nova_val = get_nova_price(code)
        if nova_val:
            df.at[index, 'nova'] = nova_val
            print(" [Nova: ✅]", end="")
        else:
            print(" [Nova: ❌]", end="")
            
        # Liloshop
        lilo_val = get_liloshop_price(code)
        if lilo_val:
            df.at[index, 'liloshop'] = lilo_val
            print(" [Lilo: ✅]")
        else:
            print(" [Lilo: ❌]")
            
        # ყოველ 5 კოდზე ვინახავთ და ვისვენებთ 1 წამით, რომ საიტმა არ დაგვბლოკოს
        if (index + 1) % 5 == 0:
            df.to_excel(output_file, index=False)
        time.sleep(1)

    # რეკომენდირებული ფასის დათვლა
    competitor_cols = ['intexgeorgia', 'bestway', 'gardex', 'nova', 'liloshop']
    df['price'] = df[competitor_cols].min(axis=1)
    df['price'] = df['price'].fillna(df['Wendershop'] if 'Wendershop' in df.columns else df['Wondershop'])

    df.to_excel(output_file, index=False)
    print(f"✅ სკრაპინგი წარმატებით დასრულდა! ფაილი მზად არის: {output_file}")

if __name__ == "__main__":
    main()
