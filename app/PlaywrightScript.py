#pip install playwright
#pip install pandas
#pip install openpyxl

#PLAYWRIGHT IMPORTS
from playwright.sync_api import sync_playwright

#TKINTER IMPORTS
import tkinter as tk
from tkinter import messagebox

#OTHER IMPORTS
import urllib, time, os, pandas as pd

class App():
    def __init__(self, excel_file, pdf_folder):

        self.excel_file = pd.read_excel(excel_file)
        self.pdf_folder = pdf_folder


    def main(self):
        with sync_playwright() as p:
            try:
                browser_path = "C:\\Program Files\\Google\\Chrome\\Application\\chrome.exe"
                browser =  p.chromium.launch(executable_path = browser_path, headless=False)

                for customers_count, name in enumerate(self.excel_file['Nome Cliente']):
                    name_str = str(name)
                    print(name, customers_count)
                    pdf_path = os.path.join(self.pdf_folder, f'{name_str}.pdf')
                    if os.path.exists(pdf_path):
                        print(f'File from {name} found!')
                        print('Loading file:', pdf_path)

                        page = browser.new_page()
                        page.goto(f'https://web.whatsapp.com/')
                        
                        page.wait_for_selector('//*[@id="app"]/div/div[2]/div[3]/header/div[2]/div/span/div[5]/div/span')

                        message_text = urllib.parse.quote(f'Olá {name}! Aqui está o seu boleto: {pdf_path}')
                        phone_number = "31999042626" #client number or array XD
                        link = f'https://web.whatsapp.com/send?phone={phone_number}&text={message_text}'
                        page.goto(link)

                        page.wait_for_selector('//*[@id="app"]/div/div[2]/div[3]/header/div[2]/div/span/div[5]/div/span')
                        page.locator('//*[@id="main"]/footer/div[1]/div/span[2]/div/div[2]/div[2]/button').click()

                        time.sleep(5)
                    else:
                        pass
            except Exception as error:
                root = tk.Tk()
                root.withdraw()
                root.lift()

                messagebox.showerror('Error: ', str(error))
                root.destroy()
                
if __name__ == '__main__':
    app_handler = App('', '') #parametros: (excel_file, pdf_folder)
    app_handler.main()