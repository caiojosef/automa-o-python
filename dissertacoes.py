from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import pyautogui
import os
import re
import pyodbc

# Função para corrigir nome do arquivo
def corrigir_nome_arquivo(nome):
    nome_corrigido = re.sub(r'[<>:"/\\|?*]', '', nome)  # Remove caracteres inválidos
    nome_corrigido = nome_corrigido.strip().lower().replace(' ', '-')  # Converte para minúsculas e substitui espaços por traços
    return nome_corrigido

# Função para clicar com JavaScript caso o elemento esteja bloqueado
def clicar_com_js(driver, elemento):
    driver.execute_script("arguments[0].click();", elemento)

# Função para enviar dados para o banco de dados
def enviar_para_banco(titulo, autor, ano, arquivo_path):
    try:
        conexao = pyodbc.connect(
            "DRIVER={ODBC Driver 11 for SQL Server};"
            "SERVER=DESKTOP-M49THPR\\SQLEXPRESS;"
            "DATABASE=DissertacoesDB;"  # Nome do banco de dados
            "UID=sa;"
            "PWD=261456;"
        )
        cursor = conexao.cursor()

        # Ler o arquivo em formato binário
        with open(arquivo_path, 'rb') as file:
            arquivo_binario = file.read()

        # Obter tipo, extensão e tamanho do arquivo
        tipo_arquivo = "application/pdf"
        extensao = ".pdf"
        tamanho_arquivo = os.path.getsize(arquivo_path)  # Tamanho do arquivo em bytes

        # Inserir os dados no banco de dados
        cursor.execute(
            """
            INSERT INTO Dissertacoes (titulo, autor, ano, arquivo, tipo_arquivo, extensao, nome_arquivo, tamanho_arquivo)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
            """,
            titulo, autor, ano, arquivo_binario, tipo_arquivo, extensao, os.path.basename(arquivo_path), tamanho_arquivo
        )
        conexao.commit()
        print(f"Dados enviados ao banco de dados: {titulo}, {autor}, {ano}, {tamanho_arquivo} bytes")
        cursor.close()
        conexao.close()
    except Exception as e:
        print(f"Erro ao enviar dados para o banco de dados: {e}")

# Função para encontrar o arquivo mais recente no diretório de downloads
def encontrar_arquivo_recente(diretorio, extensao=".pdf"):
    try:
        arquivos = [os.path.join(diretorio, f) for f in os.listdir(diretorio) if f.endswith(extensao)]
        if not arquivos:
            return None
        return max(arquivos, key=os.path.getctime)  # Retorna o arquivo mais recente
    except Exception as e:
        print(f"Erro ao procurar o arquivo no diretório {diretorio}: {e}")
        return None

# Função para buscar dados na seção de um ano específico
def buscar_dados_por_ano(driver, ano):
    try:
        wait = WebDriverWait(driver, 10)

        # Expandir a seção do ano
        try:
            botao_ano = wait.until(EC.presence_of_element_located((By.XPATH, f"//h4[@class='titulo-sanfona' and @rel='s-{ano}']")))
            clicar_com_js(driver, botao_ano)
            print(f"Botão '{ano} [+]' clicado.")
            time.sleep(2)
        except Exception as e:
            print(f"Erro ao clicar no botão '{ano} [+]': {e}")
            return

        # Localizar os itens dentro da seção do ano
        trabalhos = driver.find_elements(By.XPATH, f"//div[@class='ord-item' and ancestor::div[contains(@class, 's-{ano}')]]")

        if not trabalhos:
            print(f"Nenhum trabalho encontrado na seção {ano}.")
        else:
            for trabalho in trabalhos:
                try:
                    # Extrair título
                    titulo_elemento = trabalho.find_element(By.TAG_NAME, "h5")
                    titulo = titulo_elemento.text.strip() if titulo_elemento else ""

                    # Extrair autor
                    autor_elemento = trabalho.find_element(By.XPATH, ".//p/strong[@class='ord-indice']")
                    autor = autor_elemento.text.strip() if autor_elemento else ""

                    # Extrair ano
                    ano_elemento = trabalho.find_element(By.XPATH, ".//p/strong[not(@class)]")
                    ano = ano_elemento.text.strip() if ano_elemento else ""

                    # Verificar se os dados estão completos
                    if titulo and autor and ano:
                        print("\nDados encontrados:")
                        print(f"Título: {titulo}")
                        print(f"Autor: {autor}")
                        print(f"Ano: {ano}")

                        # Abrir o link "Veja o trabalho"
                        try:
                            link_trabalho = trabalho.find_element(By.LINK_TEXT, "Veja o trabalho")
                            clicar_com_js(driver, link_trabalho)
                            print(f"Abrindo link para o trabalho: {titulo}")
                            time.sleep(2)  # Esperar a nova aba carregar

                            # Alternar para a nova aba
                            driver.switch_to.window(driver.window_handles[-1])

                            # Pressionar Ctrl + S para abrir "Salvar como"
                            time.sleep(1)
                            pyautogui.hotkey("ctrl", "s")
                            print("Salvando o arquivo...")
                            time.sleep(2)

                            # Pressionar Enter para salvar o arquivo
                            pyautogui.press("enter")
                            time.sleep(3)  # Aguardar o download concluir

                            # Fechar a aba atual
                            driver.close()

                            # Retornar para a aba principal
                            driver.switch_to.window(driver.window_handles[0])

                            # Verificar o arquivo mais recente no diretório de downloads
                            diretorio_downloads = "C:\\Users\\caioo\\Downloads"
                            arquivo_path = encontrar_arquivo_recente(diretorio_downloads)

                            # Verificar se o arquivo foi baixado
                            if arquivo_path and os.path.exists(arquivo_path):
                                enviar_para_banco(titulo, autor, ano, arquivo_path)
                            else:
                                print(f"Arquivo não encontrado no diretório: {diretorio_downloads}")

                        except Exception as e:
                            print(f"Erro ao fazer download do trabalho {titulo}: {e}")

                except Exception as e:
                    print(f"Erro ao processar trabalho: {e}")

    except Exception as e:
        print(f"Erro ao acessar os dados da página para o ano {ano}: {e}")

# Programa principal
if __name__ == "__main__":
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=options)
    driver.get("https://www.uniara.com.br/ppg/biotecnologia-medicina-regenerativa-quimica-medicinal/repositorio-cientifico/dissertacoes/")

    print("Inicializando...")

    anos = [2024, 2023, 2022, 2021, 2020, 2019, 2018, 2017]

    for ano in anos:
        print(f"\nBuscando dados para o ano {ano}...")
        buscar_dados_por_ano(driver, ano)

    driver.quit()


