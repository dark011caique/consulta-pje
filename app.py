from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.select import Select
from time import sleep
import openpyxl

# Abrindo a planilha e a aba "processos"
planilha = openpyxl.load_workbook('dados_de_processos.xlsx')
processos = planilha['processos']  # Certifique-se de que a aba "processos" existe

# Inicializando o navegador
driver = webdriver.Chrome()
driver.get('https://pje-consulta-publica.tjmg.jus.br/')
sleep(5)

# Inserindo número da OAB
oab = driver.find_element(By.XPATH, '//*[@id="fPP:Decoration:numeroOAB"]')
sleep(1)
oab.click()
oab.send_keys('259155')

# Selecionando o estado
sleep(1)
estado = driver.find_element(By.XPATH, '//select[@id="fPP:Decoration:estadoComboOAB"]')
opcoes_uf = Select(estado)
opcoes_uf.select_by_visible_text('SP')

# Clicando no botão de pesquisa
sleep(1)
pesquisa = driver.find_element(By.XPATH, '//*[@id="fPP:searchProcessos"]')
pesquisa.click()
sleep(3)

# Capturando os links de processos
links_abrir_processo = driver.find_elements(By.XPATH, '//a[@title="Ver Detalhes"]')

# Iterando sobre os links de processos
for link in links_abrir_processo:
    janela_principal = driver.current_window_handle
    link.click()
    sleep(5)
    janelas_abertas = driver.window_handles

    for janela in janelas_abertas:
        if janela != janela_principal:
            driver.switch_to.window(janela)
            sleep(5)

            # Capturando o número do processo
            try:
                numero_processo = driver.find_element(By.XPATH, '//div[@class="propertView"]//div[@class="col-sm-12 "]').text
            except:
                numero_processo = "Não encontrado"

            # Capturando os nomes dos participantes
            nome_participante = driver.find_elements(By.XPATH, '//table[contains(@id, "processoPartesPoloAtivoResumidoList")]//span[@class="text-bold"]')
            lista_participantes = [participante.text for participante in nome_participante]

            # Adicionando os dados à planilha
            if lista_participantes:
                if len(lista_participantes) == 1:
                    processos.append(['259155', numero_processo, lista_participantes[0]])
                else:
                    processos.append(['259155', numero_processo, ', '.join(lista_participantes)])
            else:
                processos.append(['259155', numero_processo, "Sem participantes"])

            # Fechando a aba do processo
            driver.close()

    # Voltando para a janela principal
    driver.switch_to.window(janela_principal)

# Salvando a planilha
try:
    planilha.save('dados_de_processos.xlsx')
    print("Dados salvos com sucesso na planilha!")
except Exception as e:
    print(f"Erro ao salvar a planilha: {e}")

# Finalizando
driver.quit()
