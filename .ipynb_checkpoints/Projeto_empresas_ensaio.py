from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import pandas as pd
import win32com.client as win32
import pandas as pd
import time

### Configuração de navegador
servico = Service(ChromeDriverManager().install())
navegador = webdriver.Chrome(service=servico)

### Acessando o site do inmetro e buscando todas as paginas do site.

navegador.get(
    'http://www.inmetro.gov.br/laboratorios/rble/lista_laboratorios.asp?sigLab=&codLab=&tituloLab=&uf=&pais=&classe_ensaio=&area_atividade=&descr_escopo=&Submit2=Buscar')

paginas = navegador.find_elements(By.CLASS_NAME, 'menuHP')
pages = []
for pagina in paginas:
    pages.append(pagina.get_attribute('href'))

### Separando todas as paginas dos clientes de ensaio

# Tratando dados para obter apenas as paginas desejadas.
paginas = []
for page in pages:
    if "lista" in page and "pagina=" in page:
        paginas.append(page)
paginas.pop(-1)

### Encontrando os links das paginas dos perfis dos clientes e extraindo as informações desejadas


links = []
for pagina in paginas:
    navegador.get(pagina)  # acessando cada pagina onde os clientes estão
    link = navegador.find_elements(By.XPATH,
                                   '/html/body/table[3]/tbody/tr[2]/td[3]/table[2]/tbody/tr/td/table[2]/tbody')

    for l in link:
        sites = l.find_elements(By.TAG_NAME, 'A')
        for site in sites:
            if 'detalhe' in site.get_attribute('href'):
                links.append(site.get_attribute('href'))
                # o código acima pega os links dos perfi's dos clientes de cada pagina do site.

### Extraindo as informações desejadas dos clientes dos links encontrados.
empresa = []
email = []
contato = []
for link in links:
    navegador.get(link)
    if len(navegador.find_elements(By.XPATH,
                                   '/html/body/table[3]/tbody/tr[2]/td[3]/table[2]/tbody/tr/td/table/tbody/tr[7]/td[2]')) == 1:
        empresa.append(navegador.find_element(By.XPATH,
                                              '/html/body/table[3]/tbody/tr[2]/td[3]/table[2]/tbody/tr/td/table/tbody/tr[7]/td[2]').text)
    if len(navegador.find_elements(By.XPATH,
                                   '/html/body/table[3]/tbody/tr[2]/td[3]/table[2]/tbody/tr/td/table/tbody/tr[19]/td[2]/a')) == 1:
        email.append(navegador.find_element(By.XPATH,
                                            '/html/body/table[3]/tbody/tr[2]/td[3]/table[2]/tbody/tr/td/table/tbody/tr[19]/td[2]/a').text)
    else:
        email.append('')
    if len(navegador.find_elements(By.XPATH,
                                   '/html/body/table[3]/tbody/tr[2]/td[3]/table[2]/tbody/tr/td/table/tbody/tr[18]/td[2]')) == 1:
        contato.append(navegador.find_element(By.XPATH,
                                              '/html/body/table[3]/tbody/tr[2]/td[3]/table[2]/tbody/tr/td/table/tbody/tr[18]/td[2]').text)

### Colocando as informações de cada cliente em uma lista de tuplas

clientes = []
for n in range(len(empresa)):
    if email[n] == "":
        print("-|-|" * 20)
        print(f'A empresa{empresa[n]} não possui email cadastrado no site do inmetro.')
        print("-|-|" * 20)
    else:
        print("-|-|" * 10)
        print(f'A empresa {empresa[n]} representada pelo contato {contato[n]} possui o email:{email[n]}')
        clientes.append((empresa[n], email[n], contato[n]))
        print("-|-|" * 10)

### Criando planilha de excel com informações coletadas

clientes_df = pd.DataFrame.from_records(clientes)
clientes_df = clientes_df.rename(columns={0: "Empresa", 1: "E-mail", 2: "Contato"})
display(clientes_df)

clientes_df.to_excel('//servidor/Banco de Dados/clientes_ensaio.xlsx', sheet_name='Clientes_ensaio', na_rep='',
                     merge_cells=True)

### ENVIO DE E-MAIL PARA ESTABELECER CONTATO COMERCIAL
clientes_df = pd.read_excel('//servidor/Banco de Dados/Outros/clientes_ensaio.xlsx')

outlook = win32.Dispatch('outlook.application')
nao_enviado = 0
enviado = 0
for n, e in enumerate(clientes_df["E-mail"]):
    try:
        empresa = clientes_df["Empresa"][n]
        empresa = empresa.title()
        mail = outlook.CreateItem(0)
        mail.To = e
        mail.CC = ''
        mail.BCC = ''
        mail.Subject = 'Calibração de Equipamentos'
        mail.Body = f"""Olá Prezados,

Somos uma empresa de calibração de instrumentos acreditada pelo INMETRO.

Gostaríamos de estabelecer um relacionamento comercial com vocês, pensando nisso, estamos disponibilizando uma promoção de primeiro serviço, no qual o cliente recebe 10% de desconto no valor total de seu pedido. Não perca esta oportunidade e venha calibrar seus equipamentos conosco!

Para utilizar o desconto basta informar no pedido de orçamento que veio através do meu contato.

Enviei em anexo um folheto para demonstração dos itens em nosso escopo de calibrações, porém, caso necessite de mais informações consulte o nosso site:

http://www.lekas.com.br

Caso deseje um orçamento, por favor, nos solicite pelo e-mail: comercial@lekas.com.br.


Aguardamos ansiosamente o seu contato.

Juan F. C. Ladeira
Agente da Qualidade

Tridimensional Leka's Medições
Rua Rio Apa, 564 – Cordovil
(21) 3458-9449
(21) 98477-8391 (APENAS WHATSAPP)

Laboratório de Calibração acreditado pela CGCRE sob o número 71	

http://www.lekas.com.br

Esta mensagem, incluindo os seus anexos, contém informações confidenciais destinadas a indivíduo e propósito específicos, e é protegida por lei. Caso você não seja o citado indivíduo, deve apagar esta mensagem. 
É terminantemente proibida a utilização, acesso, cópia ou divulgação não autorizada das informações presentes nesta mensagem.

        """
        # Anexos (pode colocar quantos quiser):
        attachment2 = r'C:/Users/juan/Desktop/Apresentação Lekas.pdf'
        mail.Attachments.Add(attachment2)
        mail.Send()
        enviado += 1
        print(f'Email:{e} Enviado com sucesso para a {empresa}')
        print(enviado)
        sleep(45)
    except:
        print(e)
        print('Não enviou e-mail')
        nao_enviado += 1
        print(nao_enviado)
