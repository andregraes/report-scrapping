from selenium import webdriver
from selenium.webdriver.common.by import By
from docxtpl import DocxTemplate
import jinja2
import time
import random
import os
from selenium.webdriver.support.wait import WebDriverWait


student_id = "Your_name"
driver = webdriver.Edge()


urls = ["https://catalogo.anqep.gov.pt/ufcdDetalhe/5327", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5328",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5329", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5330",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5331", "https://catalogo.anqep.gov.pt/ufcdDetalhe/353677",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5333", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5335",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5336", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5337",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5338", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5341",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/353678", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5345",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/353679", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5371",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/353675", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5392",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/353681", "https://catalogo.anqep.gov.pt/ufcdDetalhe/353685",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5348", "https://catalogo.anqep.gov.pt/ufcdDetalhe/353680",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5339", "https://catalogo.anqep.gov.pt/ufcdDetalhe/4048",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5502", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5391",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5501", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5503",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5504", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5394",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5374", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5506",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5447", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5449",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5379", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5383",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5507", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5381",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5350", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5384",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5393", "https://catalogo.anqep.gov.pt/ufcdDetalhe/14349",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/14350", "https://catalogo.anqep.gov.pt/ufcdDetalhe/14351",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/14352", "https://catalogo.anqep.gov.pt/ufcdDetalhe/11719",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/12025", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5342",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5380", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5382",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/353676", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5352",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5353", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5354",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5355", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5358",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/353674", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5378",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5375", "https://catalogo.anqep.gov.pt/ufcdDetalhe/353686",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5419", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5422",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5433", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5434",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5435", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5450",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5505", "https://catalogo.anqep.gov.pt/ufcdDetalhe/5436",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/5429", "https://catalogo.anqep.gov.pt/ufcdDetalhe/12630",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/12631", "https://catalogo.anqep.gov.pt/ufcdDetalhe/12625",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/12626", "https://catalogo.anqep.gov.pt/ufcdDetalhe/12627",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/9116", "https://catalogo.anqep.gov.pt/ufcdDetalhe/9117",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/9118", "https://catalogo.anqep.gov.pt/ufcdDetalhe/9119",
        "https://catalogo.anqep.gov.pt/ufcdDetalhe/17491", "https://catalogo.anqep.gov.pt/ufcdDetalhe/17510"]

templates = ["template.docx", "template1.docx", "template2.docx", "template3.docx"]
subjects = []


def buscar_formador(parametro):
    jose_carlos = ["6011", "6012", "6013", "6015", "6016", "6017", "6018", "6021", "6024", "6025", "6026", "6072",
                   "6040",
                   "6028", "6029", "6174", "6059", "6063", "6061", "6030", "6064", "6073", "6062", "6060", "6091"]
    carlos_hipolito = ["6007", "6008", "6009", "6010", "6075", "6019", "6183", "6127", "6129"]
    ronaldo = ["6051", "6052", "6182", "6184", "6054", "6186", "6187"]
    rui_franco = ["6102"]

    if parametro in jose_carlos:
        return "José Carlos"
    elif parametro in carlos_hipolito:
        return "Carlos Ferreira"
    elif parametro in ronaldo:
        return "Ronaldo Carvalho"
    elif parametro in rui_franco:
        return "Rui Franco"
    else:
        return "Não encontrado em nenhuma lista"


def template_selection(list):
    random.shuffle(list)

    random_index = random.randint(0, len(list) - 1)

    selected_item = list[random_index]

    return selected_item


for item in urls:
    driver.implicitly_wait(5)
    driver.get(item)
    texto = template_selection(templates)
    tpl = DocxTemplate(texto)
    temp_objectives = []
    objectives = driver.find_element(By.CSS_SELECTOR,
                                    "#search-details > div > div > div > div:nth-child(3) > div > div > div > div > ul").text.split(
        "\n")

    temp_content = []
    content = driver.find_element(By.CSS_SELECTOR,
                                   "#search-details > div > div > div > div:nth-child(4) > div > div > div > div > ul").text.split(
        "\n")

    driver.implicitly_wait(5)
    for obj in objectives:
        temp_objectives.append(obj)

    for cont in content:
        temp_content.append(cont)

    data = {
        'ufcd': driver.find_element(By.CSS_SELECTOR,
                                    "#search-details > div > div > div > div.innerpage-heading > h1").text,
        'codigo_ufcd': driver.find_element(By.CSS_SELECTOR,
                                           "#search-details > div > div > div > div:nth-child(2) > div > div > div > div > div > div > div > p:nth-child(4)").text.split(
            ":")[1].lstrip(),
        'formador': buscar_formador(driver.find_element(By.CSS_SELECTOR,
                                                        "#search-details > div > div > div > div:nth-child(2) > div > div > div > div > div > div > div > p:nth-child(4)").text.split(
            ":")[1].lstrip()),
        'objetivos': temp_objectives,
        'conteudo': temp_content
    }
    context = {
        'ufcd': data['ufcd'],
        'codigo_ufcd': data['codigo_ufcd'],
        'formador': data['formador'],
        'objetivos': data['objetivos'],
        'conteudo': data['conteudo']
    }
    print(data)
    # print(context)

    jinja_env = jinja2.Environment(autoescape=True)
    tpl.render(context, jinja_env)

    output_directory = 'C:/new_reflexoes/' + context['codigo_ufcd'] + '_' + context['ufcd']
    if not os.path.exists(output_directory):
        os.makedirs(output_directory)

    tpl.save(output_directory + '/PRA_{0}{1}.docx'.format(context['codigo_ufcd'], student_id))
    time.sleep(3)
    subjects.append(data)
