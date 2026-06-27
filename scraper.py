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
    """ბრაუზერის ოპტიმიზირებული კონფიგურაცია GitHub Actions-ისთვის"""
    options = webdriver.ChromeOptions()
    options.add_argument("--headless=new") # ახალი, უფრო სტაბილური headless რეჟიმი
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage") # 🛠️ ხელს უშლის მეხსიერების გამო გაყინვას Linux-ზე
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--window-size=1920,1080")
    
    # მომხმარებლის იმიტაცია
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    
    driver.set_page_load_timeout(20) # დროის ლიმიტი გვერდისთვის
    return driver

def get_nova_price(driver, code):
    try:
        driver.get("https://nova.ge/")
        time.sleep(1.5)
        
        search_box = WebDriverWait(driver, 7).until(
            EC.presence_of_element_located((By.ID, "search_input"))
        )
        search_box.clear()
        search_box.send_keys(str(code))
        search_box.send_keys(Keys.RETURN)
        
        # ველოდებით კონკრეტულად ფასის გამოჩენას
        price_element = WebDriverWait(driver, 7).until(
            EC.presence_of_element_located((By.CLASS_NAME, "ty-price-num"))
        )
        return float(price_element.text.replace("₾", "").replace(",", "").strip())
    except Exception as e:
        # არ აგდებს ერორს, უბრალოდ აგრძელებს მუშაობას
        return None

def get_liloshop_price(driver, code):
    try:
        driver.get("https://liloshop.ge/")
        time.sleep(1.5)
        
        search_box = WebDriverWait(driver, 7).until(
            EC.presence_of_element_located((By.NAME, "q"))
        )
        search_box.clear()
        search_box.send_keys(str(code))
        search_box.send_keys(Keys.RETURN)
        
        price_element = WebDriverWait(driver, 7).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, ".product-price, .price"))
        )
        return float(price_element.text.replace("₾", "").replace(" ", "").replace(",", "").strip())
    except Exception as e:
        return None

def main():
    input_file = "ai.xlsx"
    output_file = "fasebi_ganaxlebuli.xlsx"
    
    if not os.path.exists(input_file):
        print(f"⚠️ ფაილი {input_file} ვერ მოიძებნა.")
        return

    df = pd.read_excel(input_file)
    driver = init_driver()
    
    print(f"🚀 სკრაპინგი დაიწყო. სულ დასამუშავებელია {len(df)} კოდი...")
    
    for index, row in df.iterrows():
        code = str(row['code'])
        print(f"🔄 [{index + 1}/{len(df)}] ვამუშავებ კოდს: {code}...", end="", flush=True)
        
        # Nova
        nova_val = get_nova_price(driver, code)
        if nova_val:
            df.at[index, 'nova'] = nova_val
            print(" [Nova: ✅]", end="")
        else:
            print(" [Nova: ❌]", end="")
            
        # Liloshop
        lilo_val = get_liloshop_price(driver, code)
        if lilo_val:
            df.at[index, 'liloshop'] = lilo_val
            print(" [Lilo: ✅]")
        else:
            print(" [Lilo: ❌]")
            
        # პერიოდულად (ყოველ 5 პოზიციაზე) ვინახავთ ფაილს, რომ შუა გზაში გათიშვისას მონაცემები არ დაიკარგოს
        if (index + 1) % 5 == 0:
            df.to_excel(output_file, index=False)

    # საბოლოო ფასის გამოთვლა
    competitor_cols = ['intexgeorgia', 'bestway', 'gardex', 'nova', 'liloshop']
    df['price'] = df[competitor_cols].min(axis=1)
    df['price'] = df['price'].fillna(df['Wondershop'])

    df.to_excel(output_file, index=False)
    driver.quit()
    print(f"✅ სკრაპინგი დასრულდა წარმატებით! ფაილი შენახულია: {output_file}")

if __name__ == "__main__":
    main()
