import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager

def init_driver():
    """ბრაუზერის კონფიგურაცია ფონურ რეჟიმში"""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

def get_nova_price(driver, code):
    """ფასის წამოღება nova.ge-დან"""
    try:
        url = f"https://nova.ge/?subcats=Y&pcode_from_q=Y&pshort=Y&pfull=Y&pname=Y&pkeywords=Y&search_performed=Y&q={code}&dispatch=products.search"
        driver.get(url)
        time.sleep(3)
        price_element = driver.find_element(By.CLASS_NAME, "ty-price-num")
        return float(price_element.text.replace("₾", "").replace(",", "").strip())
    except Exception:
        return None

def get_liloshop_price(driver, code):
    """ფასის წამოღება liloshop.ge-დან"""
    try:
        url = f"https://liloshop.ge/search?q={code}"
        driver.get(url)
        time.sleep(3)
        price_element = driver.find_element(By.CSS_SELECTOR, ".product-price, .price")
        return float(price_element.text.replace("₾", "").replace(" ", "").replace(",", "").strip())
    except Exception:
        return None

def main():
    input_file = "ai.xlsx"
    output_file = "fasebi_ganaxlebuli.xlsx"
    
    # თუ ფაილი არ არსებობს, ვქმნით სატესტო ვერსიას
    if not os.path.exists(input_file):
        print(f"⚠️ ფაილი {input_file} ვერ მოიძებნა. იქმნება სატესტო შაბლონი...")
        demo_data = {
            'code': [68672, 68676, 56182],
            'Wondershop': [4, 5, 125],
            'intexgeorgia': [None, None, None],
            'bestway': [None, None, None],
            'gardex': [None, None, None],
            'nova': [None, None, None],
            'liloshop': [None, None, None],
            'price': [None, None, None]
        }
        pd.DataFrame(demo_data).to_excel(input_file, index=False)

    df = pd.read_excel(input_file)
    driver = init_driver()
    
    print("🚀 სკრაპინგი დაიწყო...")
    
    for index, row in df.iterrows():
        code = str(row['code'])
        print(f"🔄 მუშავდება კოდი: {code}")
        
        # Nova
        nova_val = get_nova_price(driver, code)
        if nova_val:
            df.at[index, 'nova'] = nova_val
            
        # Liloshop
        lilo_val = get_liloshop_price(driver, code)
        if lilo_val:
            df.at[index, 'liloshop'] = lilo_val
            
        # აქ შეგიძლიათ ჩაამატოთ სხვა საიტების ფუნქციებიც...

    # რეკომენდირებული ფასის გამოთვლა (მინიმალური კონკურენტებს შორის)
    competitor_cols = ['intexgeorgia', 'bestway', 'gardex', 'nova', 'liloshop']
    df['price'] = df[competitor_cols].min(axis=1)
    df['price'] = df['price'].fillna(df['Wendershop'] if 'Wendershop' in df.columns else df['Wondershop'])

    # შენახვა
    df.to_excel(output_file, index=False)
    driver.quit()
    print(f"✅ სამუშაო დასრულდა! შედეგი შენახულია: {output_file}")

if __name__ == "__main__":
    main()
