import os
import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

def init_driver():
    """ბრაუზერის კონფიგურაცია ფონურ რეჟიმში"""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--window-size=1920,1080")
    # იმიტაცია, რომ ნამდვილი ბრაუზერია და არა ბოტი
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    service = Service(ChromeDriverManager().install())
    return webdriver.Chrome(service=service, options=options)

def get_nova_price(driver, code):
    """ხსნის მთავარ გვერდს, წერს კოდს საძიებო ველში და ეძებს"""
    try:
        driver.get("https://nova.ge/")
        time.sleep(2)
        
        # ვპოულობთ საძიებო ველს ID-ით ან კლასით (ნოვას საიტზე ხშირად 'search_input' ჰქვია)
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "search_input"))
        )
        
        search_box.clear()
        search_box.send_keys(str(code))
        search_box.send_keys(Keys.RETURN) # აჭერს Enter-ს
        
        time.sleep(3) # ველოდებით შედეგების ჩატვირთვას
        
        # ფასის ელემენტის აღება
        price_element = driver.find_element(By.CLASS_NAME, "ty-price-num")
        price_text = price_element.text.replace("₾", "").replace(",", "").strip()
        return float(price_text)
    except Exception as e:
        print(f"❌ Nova-ზე კოდი {code} ვერ მოიძებნა ან მოხდა შეცდომა")
        return None

def get_liloshop_price(driver, code):
    """ხსნის ლილოშოპის მთავარს, პოულობს საძიებო გრაფას და ეძებს"""
    try:
        driver.get("https://liloshop.ge/")
        time.sleep(2)
        
        # ლილოშოპის საძიებო ველის პოვნა (სახელით 'q')
        search_box = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.NAME, "q"))
        )
        
        search_box.clear()
        search_box.send_keys(str(code))
        search_box.send_keys(Keys.RETURN)
        
        time.sleep(3)
        
        price_element = driver.find_element(By.CSS_SELECTOR, ".product-price, .price")
        price_text = price_element.text.replace("₾", "").replace(" ", "").replace(",", "").strip()
        return float(price_text)
    except Exception as e:
        print(f"❌ Liloshop-ზე კოდი {code} ვერ მოიძებნა ან მოხდა შეცდომა")
        return None

def main():
    input_file = "ai.xlsx"
    output_file = "fasebi_ganaxlebuli.xlsx"
    
    if not os.path.exists(input_file):
        print(f"⚠️ ფაილი {input_file} ვერ მოიძებნა.")
        return

    df = pd.read_excel(input_file)
    driver = init_driver()
    
    print("🚀 სკრაპინგი დაიწყო რეალური ძებნის რეჟიმით...")
    
    for index, row in df.iterrows():
        code = str(row['code'])
        print(f"🔄 ძებნაშია კოდი: {code}")
        
        # Nova ძებნა
        nova_val = get_nova_price(driver, code)
        if nova_val:
            df.at[index, 'nova'] = nova_val
            
        # Liloshop ძებნა
        lilo_val = get_liloshop_price(driver, code)
        if lilo_val:
            df.at[index, 'liloshop'] = lilo_val
            
    # რეკომენდირებული ფასის დათვლა
    competitor_cols = ['intexgeorgia', 'bestway', 'gardex', 'nova', 'liloshop']
    df['price'] = df[competitor_cols].min(axis=1)
    df['price'] = df['price'].fillna(df['Wondershop'])

    df.to_excel(output_file, index=False)
    driver.quit()
    print(f"✅ სამუშაო დასრულდა! ფაილი შენახულია: {output_file}")

if __name__ == "__main__":
    main()
