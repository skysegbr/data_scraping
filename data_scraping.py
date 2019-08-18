
# Script Python
# Data Scraping
# 18-08-2019

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
#import pdb
from bs4 import BeautifulSoup
import xlsxwriter

# Grava os dados em arquivo XLSX
def grava_dados(lista, path_nome, cabecalho):
    workbook = xlsxwriter.Workbook(path_nome)
    worksheet = workbook.add_worksheet()
    worksheet.set_column('A:A', 20)
    bold = workbook.add_format({'bold': True})
    for col, t in enumerate(cabecalho):
        worksheet.write(0, col, t, bold)
        col += 1


    for idx, val in enumerate(lista):
        worksheet.write(idx + 1, 0, val['NOME_PRODUTO'])
        worksheet.write(idx + 1, 1, val['COD_PRODUTO'])
        worksheet.write(idx + 1, 2, val['PRECO_PRODUTO'])
        worksheet.write(idx + 1, 3, val['DESC_PRODUTO'])

    workbook.close()



# Inicializa o driver de busca
def init_driver_busca(url, tag, busca):
    driver = webdriver.Chrome('./chromedriver')
    driver.get(url)
    if busca != None:
        pesquisa = driver.find_element_by_id(tag)
        
        pesquisa.clear()
        pesquisa.send_keys(busca)
        time.sleep(1)
        pesquisa.send_keys(Keys.RETURN)
        time.sleep(2)
    return driver


# Retorna informações do produto utilizando as tags html para busca
def ret_info_prod(driver, item, tag, tag2, tag3, param_name):
    html = driver.page_source
    bs = BeautifulSoup(html, 'html.parser')
    f = bs.find_all(tag, {'class': param_name})
    
    #print(f[item])    
    if tag2 != None:
        if tag2 == 'value':
            produto = f[item][tag2]  
        elif tag3 == 'href':
            produto = f[item].find('a', href=True)[tag3]
        elif tag2 == 'p':
            produto = f[item].find_all(tag2)
        elif tag2 == 'id':
            produto = f[item].find_all(tag2)
        else:
            produto = f[item].find(tag2).text
    else:
        produto = f[item].text
    return produto


# Main - inicialização do script
print('Data Scraping iniciado')
url = "https://buscando2.extra.com.br"
driver = init_driver_busca(url, "strBusca", "petshop")

data_prod = {}
data_lis = []

for item in range(20):
    try:        

        url = ret_info_prod(driver, item, 'div', 'a', 'href', 'nm-product-name')
        if 'https:' not in url:
            url = 'https:' + url

        #print(url)
        driver2 = init_driver_busca(url, None, None)

        descricao = ret_info_prod(driver2, 0, 'div', None, None, 'descricao')
        #descricao = ret_info_prod(driver2, 1, 'label', 'id', None, 'ctl00_Conteudo_ctl45_DetalhesProduto_lblTitulo')

        data_prod = {}
        data_prod.update({'NOME_PRODUTO': ret_info_prod(driver, item, 'div', 'a', None, 'nm-product-name')})
        data_prod.update({'PRECO_PRODUTO': ret_info_prod(driver, item, 'span', None, None, 'nm-price-value').strip()})
        data_prod.update({'COD_PRODUTO': ret_info_prod(driver, item, 'div', "value", None, 'yv-review-quickreview')})        
        data_prod.update({'DESC_PRODUTO': descricao.strip().strip('\n')})
        data_lis.append(data_prod)
        #print(data_prod)
        #print(data_lis)
        
        #print("\n")

        #time.sleep(4)
        driver2.close()
    except Exception as e:
        print(e)
        driver2.close()

# Grava em xmlx
cabecalho = ['NOME_PRODUTO', 'COD_PRODUTO', 'PRECO_PRODUTO', 'DESC_PRODUTO']
grava_dados(data_lis, 'pet_shop_itens.xlsx', cabecalho)

driver.close()

# Fim do script
print('Data Scraping finalizado')
