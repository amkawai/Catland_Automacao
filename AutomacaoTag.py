from selenium import webdriver
from selenium.webdriver.common.keys import Keys
import time
import sys

navegador = webdriver.Firefox()

navegador.get('https://app.simples.vet/login/login.php')

navegador.find_element('id','l_usu_var_email').send_keys(sys.argv[1])
navegador.find_element('id','l_usu_var_senha').send_keys(sys.argv[2])
navegador.find_element('id','l_usu_var_senha').send_keys(Keys.ENTER)
time.sleep(1.5)


def Atualiza_tag (num_ani, lista_tag):
 navegador.get("https://app.simples.vet/principal/cliente/cliente.php")

 navegador.find_element("id", "p__ani_var_nome").send_keys(num_ani)
 navegador.find_element("id", "p__ani_var_nome").send_keys(Keys.ENTER)

 time.sleep(1)
 # caso o nome da classe tenha espaço, o código find_element("class") se confunde
 # entao é necessario usar o css selector, colocando ponto na frente do nome e tb
 # no lugar dos espaços do nome da classe...
 navegador.find_element("css selector",".linkAnimalLista.animalMarcado").click()

 time.sleep(1)
 navegador.find_element("id", "divDadosAnimal").click()

 ## clica na aba Extras, depois do nome
 time.sleep(1.5)
 navegador.find_element("css selector", "[href='#tabExtrasAni']").click()

 time.sleep(3)
 ## cadastra as tags presentes em lista_tag
 for i in range(0, len(lista_tag)):
  navegador.find_element("id","ani_txt_tag_tag").send_keys(lista_tag[i])
  navegador.find_element("id","ani_txt_tag_tag").send_keys(Keys.ENTER)
  time.sleep(0.5)

 time.sleep(3)
 ## botao de salvar
 navegador.find_element("id","btn_salvar_ani").click()

with open("codigos") as file:
 content = file.read().splitlines() 

for linha in content:
 Atualiza_tag(linha, ['Disponível'])
