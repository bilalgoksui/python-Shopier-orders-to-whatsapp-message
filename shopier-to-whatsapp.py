import tkinter as tk
from tkinter import ttk
import requests
import pandas as pd
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
from time import sleep
import pandas
import time
from selenium.webdriver.common.keys import Keys
from selenium.webdriver import ActionChains

new_data = []

def wpbot(api_key_check):

    excel_file = f"{api_key_check}_unfulfilled_orders.xlsx"
    excel_data = pandas.read_excel(excel_file, sheet_name='Sheet1')
    count = 0
    message = entry_message.get('1.0', 'end-1c')


        
    driver = webdriver.Chrome(ChromeDriverManager().install())
    driver.get('https://web.whatsapp.com')
    time.sleep(45)
    # input("Press ENTER after login into Whatsapp Web and your chats are visiable.")
    for column in excel_data['Phone'].tolist():
        try:
            #message ekstra eklenecek
            url = 'https://web.whatsapp.com/send?phone=' + str(excel_data['Phone'][count]) + '&text=' + message
            sent = True
            driver.get(url)
            try:
                time.sleep(10)

                mesaj = driver.find_element(By.XPATH ,"//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]")
                mesaj.click()
                time.sleep(10)
                butn = driver.find_element(By.XPATH ,"//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[2]/button")
                butn.click()
                sleep(10)
                actions = ActionChains(driver)
                actions.send_keys(Keys.ENTER)
                actions.perform()
            except Exception as e:
                print("Sorry message could not sent to " + str(excel_data['Phone'][count]))
                new_data.append([excel_data['Name'][count], excel_data['Email'][count], excel_data['Phone'][count], excel_data['Price'][count], excel_data['Date Created'][count]])

            else:
                # time.sleep(2)
                # mesaj = driver.find_element(By.XPATH ,"//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[1]/div/div[1]")
                # mesaj.click()
                # time.sleep(1)
                # butn = driver.find_element(By.XPATH ,"//*[@id='main']/footer/div[1]/div/span[2]/div/div[2]/div[2]/button")
                # butn.click()
                sleep(3)
                print('Message sent to: ' + str(excel_data['Phone'][count]))
            count = count + 1
        except Exception as e:
            print('Failed to send message to ' + str(excel_data['Phone'][count]) + str(e))
    driver.quit()
    print("The script executed successfully.")
    new_df = pd.DataFrame(new_data, columns=["Name", "Email", "Phone", "Price", "Date Created"])
    excel_file = f"{api_key_check}_failed_messages.xlsx"
    new_df.to_excel(excel_file, index=False)

def fetch_orders_dynamic(api_key_check):
    api_key="Bearer ..----"


   
    url = "https://api.shopier.com/v1/orders"

    headers = {
        "accept": "application/json",
        "authorization": api_key
    }

    all_unfulfilled_orders = []  # Tüm sayfalardaki verileri toplamak için boş bir liste
    date_start = "2023-04-23"
    time_start = "13:24:51+0300"
    time = datetime.now().strftime("%Y-%m-%dT%H:%M:%S+0300")
    for page in range(1, 6):  # İlk 5 sayfayı döngüyle gez
        params = {
            "dateStart": f"{date_start}T{time_start}",
            "dateEnd": f"{time}",
            "fulfillmentStatus": "unfulfilled",
            "limit": 50,
            "page": page,  
            "sort": "dateDesc"
        }

        response = requests.get(url, params=params, headers=headers)

        if response.status_code == 200:
            orders = response.json()
            unfulfilled_orders = [order for order in orders if order["status"] == "unfulfilled"]
            all_unfulfilled_orders.extend(unfulfilled_orders)  # Sayfadaki verileri listeye ekle
        else:
            print(f"Error: {response.status_code} - {response.text}")
            break  # Hata durumunda döngüyü kır

    print(f"Found {len(all_unfulfilled_orders)} unfulfilled orders in first 5 pages.")


    try:
        excel_file = f"{api_key_check}_unfulfilled_orders.xlsx"

        df = pd.read_excel(excel_file)
    except FileNotFoundError:
        df = pd.DataFrame()

    new_data = []
    for order in all_unfulfilled_orders:
        name = order['shippingInfo']['firstName'] + ' ' + order['shippingInfo']['lastName']
        email = order['shippingInfo']['email']
        phone = order['shippingInfo']['phone']
        price = order['totals']['total']
        date_created = order['dateCreated']
        new_data.append([name, email, phone, price, date_created])
    new_df = pd.DataFrame(new_data, columns=["Name", "Email", "Phone", "Price", "Date Created"])
    df = pd.concat([df, new_df], ignore_index=True)
    excel_file = f"{api_key_check}_unfulfilled_orders.xlsx"
    df.to_excel(excel_file, index=False)


def fetch_orders():
    fetch_orders_dynamic('*')





root = tk.Tk()
root.title("Shoper 2 Whatsapp BOT")
root.configure(bg="#66347F")  
root.iconbitmap("*.ico")

root.geometry("800x800")

shopier_label = tk.Label(root, text="Get Orders from Shopier")
shopier_label.pack()
ttk.Separator(root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)

bot_button = tk.Button(root,bg="#A4DE02",  text="Shopier Orders", command=fetch_orders)
bot_button.pack()

ttk.Separator(root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)

wp_label = tk.Label(root,bg="green",  text="Whatsapp  Massege Bot")
wp_label.pack()

label_message = ttk.Label(root, text="Message:")
label_message.pack()
entry_message = tk.Text(root, height=10)
entry_message.pack()

ttk.Separator(root, orient=tk.HORIZONTAL).pack(fill=tk.X, padx=10, pady=5)

wp_button = tk.Button(root,bg="#A4DE02", text="Start Sending  ", command=wpbot(),)
wp_button.pack()



root.mainloop()
