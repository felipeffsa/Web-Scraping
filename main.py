import os
import threading
import time
import tkinter as tk

from selenium import webdriver
from selenium.common.exceptions import (ElementClickInterceptedException,
                                        NoSuchElementException)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

from arquivo_openpy import FileOpenpy

fileopen = FileOpenpy()


class Scraping:

    

    def iniciar(self):
        lugar = lugar_input.get()
        self.chromeConfig()
        self.chromeLinks()
        self.loopLogical(lugar)
        fileopen.openDiretorio(lugar)

        self.driver.quit()

    def tread(self):
        thead = threading.Thread(target=self.iniciar)
        thead.start()
        janela.withdraw()
        self.nova_janela()


    def chromeConfig(self):
        self.chorme_options = Options()
        self.chorme_options.add_argument('--headless')
        self.chorme_options.add_argument('--disable-gpu')
        self.driver = webdriver.Chrome(options=self.chorme_options)
        self.driver.set_window_size(width=400, height=800)
        self.driver.get("https://www.airbnb.com.br/")
    


    def chromeLinks(self):
        pesquisar = lugar_input.get()
        time.sleep(5)
        self.ok_cockie = self.driver.find_element('xpath', '/html/body/div[5]/div/div/div[1]/div/div[4]/section/div[2]/div[2]/button').click()
        time.sleep(2)
        self.aceitar = self.driver.find_element('xpath', '/html/body/div[11]/div/section/div/div/div[2]/div/div[3]/div/div/div/div/div/div/div/div[2]/div/div/div[2]/div/div/div/span/button').click()
        time.sleep(2)
        lupa_url = '/html/body/div[5]/div/div/div[1]/div/div[3]/div/div[1]/div[1]/div[1]/div/div[1]/div/div/div/div[1]/div/button/div[1]'
        self.lupa = self.driver.find_element('xpath', lupa_url).click()
        time.sleep(2)
        self.outra_lupa = self.driver.find_element('xpath', '//*[@id="accordion-body-/homes-1"]/section/div[1]/div[1]/div/div/form/div/div/label/div[1]').click()
        self.input_pesquisa = self.driver.find_element('xpath', '//*[@id="/homes-1-input"]').send_keys(pesquisar)
        self.input_ok = self.driver.find_element('xpath', '//*[@id="/homes-1-input"]').send_keys(Keys.ENTER)
        time.sleep(2)
        self.pular_calendario = self.driver.find_element('xpath', '//*[@id="accordion-body-/homes-2"]/section/div/footer/button[1]').click()
        time.sleep(2)
        self.adulto = self.driver.find_element('xpath', '//*[@id="stepper-adults"]/button[2]/span').click()
        time.sleep(2)
        self.adulto_pesquisar = self.driver.find_element('xpath', '//*[@id="vertical-tabs"]/div[3]/footer/button[2]/span[1]/span').click()
        time.sleep(5)
    
    def nova_janela(self):

        self.nova_open = tk.Tk()
        self.nova_open.title("Programa em execução")
        self.nova_open.geometry("400x300+500+500")
        label_nova = tk.Label(self.nova_open, text="Por favor aguarde", font=12)
        label_baixo = tk.Label(self.nova_open, text="A operação pode levar 3 minutos", font=12)
        label_nova.grid(column=1, row=0)
        label_baixo.grid(column=1, row=1)
        self.nova_open.mainloop()


    def loopLogical(self, mostrar):
        fileopen.criar()
        numero = 2
    
        
        
        while True:
            lugares = self.driver.find_elements(By.XPATH, '//div[contains(@class, "t1jojoys")]')
            estrelas = self.driver.find_elements(By.XPATH, '//span[contains(@class, "t1a9j9y7")]')
            preços = self.driver.find_elements(By.XPATH, '//div[contains(@style, "--pricing-guest-primary-line-font-size: 0.9375rem")]')

            for lugar, estrela, preço in zip(lugares, estrelas, preços):
                print(lugar.text)
                print(estrela.text)
                print(preço.text.replace('noite', ""))
                fileopen.openpy_laço(numero, lugar, estrela, preço)
                numero += 1
        
            time.sleep(4)
            
            try:
                
                self.element = self.driver.find_element('xpath', '//*[@id="site-content"]/div/div[2]/div/div/div/div[2]/a')
                self.link = self.element.get_attribute("href")
                self.driver.get(self.link)
                time.sleep(5)

            except ElementClickInterceptedException:
            
                time.sleep(2)
                self.element.click()
            
            except NoSuchElementException:
                fileopen.openDiretorio(mostrar)
                if var.get():
                    nome_arquivo = mostrar
                    nome_excel = nome_arquivo + ".xlsx"
                    link = r"C:\Users\felip\Desktop\\{}".format(nome_excel)
                    os.startfile(link)
                else:
                    pass
                print("Nao existe mais elemento na página")
                janela.destroy()
                self.nova_open.destroy()
                break
            
            
        


a = Scraping()




# TK INTER CONFIG ---------------------------------------------------------------------

def nova_janela():
    nova_open = tk.Tk()
    nova_open.title("Programa em execução")
    nova_open.geometry("400x300+500+500")
    label_nova = tk.Label(nova_open, text="Por favor aguarde", font=12)
    label_baixo = tk.Label(nova_open, text="A operação pode levar 3 minutos", font=12)
    label_nova.grid(column=1, row=0)
    label_baixo.grid(column=1, row=1)
    nova_open.mainloop()



janela = tk.Tk()
janela.title("AirBnb Scraping")
janela.configure(bg='#abdbe3')
#modifica o tamanho da tela

janela.minsize(400, 200)
janela.resizable(False, False)
janela.iconbitmap(r"C:\Users\felip\OneDrive\Área de Trabalho\hotel_icon.ico")
#--------------------------------------------------
rotulo = tk.Label(janela, text="Digite um lugar aqui", bg='#abdbe3', font=('Arial', 18), padx=100, pady=15)
rotulo.grid(column=0, row=0)
var = tk.BooleanVar()
lugar_input = tk.Entry(janela, width=40)
lugar_input.grid(column=0, row=3)
print(lugar_input.get())

#Mudando o tamanho da fonte

abrir_excel = tk.Checkbutton(janela, text="Abrir arquivo ao terminar", bg='#abdbe3', font=('Arial', 10), variable=var)
abrir_excel.grid(column=0, row=4, pady=10)
botao = tk.Button(janela, text="OK", width=7, font=10, command=a.tread)
botao.grid(column=0, row=8)
feito_por = tk.Label(janela, text="By: Felipe Fernandes",bg='#abdbe3')
feito_por.grid(column=0, row=9, pady=30)


janela.mainloop()