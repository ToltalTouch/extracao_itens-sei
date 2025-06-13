import logging
import os
import threading
import time
import pandas as pd
from PIL import Image
import pytesseract
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException, NoSuchElementException, TimeoutException

from login_sei import SeiLogin
from login_sei import PromptWindow

atual_dir = os.path.dirname(os.path.abspath(__file__))

logging.basicConfig(
    filename=os.path.join(atual_dir, 'extracao_itens_sei.log'),
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s'
)

sei_url = "http://sei.antt.gov.br/"

excel_path = os.path.join(atual_dir, 'excel', 'itens_extraidos.xlsx')
chromedriver_path = os.path.join(atual_dir, 'chromedriver-win64', 'chromedriver.exe')

screenshot_path = os.path.join(atual_dir, 'screenshots')
if not os.path.exists(screenshot_path):
    os.makedirs(screenshot_path)

download_dir = os.path.join(atual_dir, 'downloads')
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

df = pd.read_excel(excel_path, sheet_name='PROCESSO').dropna(subset=['PROCESSO'])
df = df[df['PROCESSO'].str.strip() != '']  # Remove strings vazias ou com espaços
process_numbers = df['PROCESSO'].tolist()

colunas = ["PROCESSO", "NOME ARQUIVO", "NOME", "MATERIAL", "MODELO", "TAMANHO/GENERO", "QUANTIDADE"]

chrome_options = Options()

def encontrar_frame(frame_id):
    try:
        WebDriverWait(sei.driver, 10).until(
            lambda d: d.find_element(By.ID, frame_id) if d.find_elements(By.ID, frame_id) else False
        )
        if sei.driver.find_elements(By.ID, frame_id):
            sei.driver.switch_to.frame(frame_id)
            logging.info(f"Frame com id {frame_id} encontrado e selecionado.")
    except TimeoutException:
        logging.error(f"Frame com id {frame_id} não encontrado.")

#def config_ocr():


def screenshot(process_number, documento_titulo):
    try:
        logging.info("Localizando tabela com informações do servidor")
        tabela_servidor = sei.driver.find_element(By.XPATH, "/html/body/table[1]")

        sei.driver.execute_script("arguments[0].scrollIntoView();", tabela_servidor)
        try:
            tabela_servidor.screenshot(screenshot_path + f"/screenshot_{process_number}_{documento_titulo}_servidor.png")
            logging.info(f"Screenshot do servidor salvo: {screenshot_path}/screenshot_{process_number}_{documento_titulo}_servidor.png")
        except WebDriverException as e:
            logging.error(f"Erro ao salvar screenshot do servidor: {e}")
            
        except (Exception, WebDriverException) as e:
            logging.error(f"Erro ao localizar e realizar screenshot da tabela do servidor: {e}")
            
        try:
            logging.info("Localizando tabela com informações do documento")
            tabela_itens = sei.driver.find_element(By.XPATH, "/html/body/table[2]/tbody")
            try:
                sei.driver.execute_script("arguments[0].scrollIntoView();", tabela_itens)
                tabela_itens.screenshot(screenshot_path + f"/screenshot_itens_{process_number}_{documento_titulo}_itens.png")
                logging.info(f"Screenshot dos itens salvo: {screenshot_path}/screenshot_itens_{process_number}_{documento_titulo}_itens.png")
            except WebDriverException as e:
                logging.error(f"Erro ao salvar screenshot dos itens: {e}")
                
        except (Exception, WebDriverException) as e:
            logging.error(f"Erro ao localizar e realizar screenshot da tabela dos itens: {e}")
            
    except (Exception, WebDriverException) as e:
        logging.error(f"Erro ao capturar screenshot: {e}")

def encontrar_arquivos(process_number):
    try:
        for process_number in process_numbers:
            logging.info(f"Iniciando busca para o processo: {process_number}")
            WebDriverWait(sei.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'txtPesquisaRapida'))
            ).send_keys(process_number + Keys.RETURN)

            try:
                encontrar_frame('ifrArvore')
                logging.info("Frame ifrArvore selecionado.")
            except Exception as e:
                logging.error(f"Erro ao selecionar frame ifrArvore: {e}")
                continue

            try:
                frm_arvore = WebDriverWait(sei.driver, 10).until(
                    EC.presence_of_element_located((By.ID, 'frmArvore'))
                )
                logging.info("frmArvore localizado.")
                itens_form = frm_arvore.find_elements(By.TAG_NAME, 'a')
            except Exception as e:
                logging.error(f"Erro ao localizar frmArvore ou seus itens: {e}")
                # Salva o HTML para depuração
                with open(f"erro_{process_number}.html", "w", encoding="utf-8") as f:
                    f.write(sei.driver.page_source)
                continue

            termos_encontrados = [item for item in itens_form if item.text.strip().lower().startswith("termo recebimento")]

            if termos_encontrados:
                
                logging.info(f"Processando termo: {len(termos_encontrados)} para o processo {process_number}")

                for idx, item_form in enumerate(termos_encontrados):
                    try:
                        documento_titulo = item_form.text.strip()
                        logging.info(f"Processando documento: {idx+1}/{len(termos_encontrados)} - {documento_titulo}")
                        df.loc[df["NOME ARQUIVO"] == documento_titulo, "PROCESSO"] = process_number
                                
                        item_form.click()
                        logging.info(f"Documento {documento_titulo} encontrado e selecionado.")

                    except Exception as e:
                        logging.error(f"Erro ao clicar no item {idx+1} do processo {process_number}: {e}")
                        continue
                        
                    sei.driver.switch_to.default_content()
                        
                    try:
                        logging.info("Encontrado o frame para screenshot.")
                        encontrar_frame('ifrVisualizacao')
                        encontrar_frame('ifrArvoreHtml')

                        try:
                            screenshot(process_number, documento_titulo)
                        except Exception as e:
                            logging.error(f"Erro ao capturar scrrenshot para {process_number}_{documento_titulo}")

                    except Exception as e:
                        logging.error(f"Erro ao encontrar o frame para screenshot: {e}")
                        continue
                        
                    sei.driver.switch_to.default_content()
                else:
                    logging.info(f"Nenhum termo encontrado para o processo {process_number}")

    except (Exception, TimeoutException, NoSuchElementException) as e:
        logging.error(f"Erro ao processar o processo {process_number}: {e}")


def main():
    WebDriverWait(sei.driver, 240).until(
        EC.presence_of_element_located((By.XPATH, '//*[@id="divInfraAreaTela"]'))
    )
    encontrar_arquivos()

if __name__ == "__main__":
    sei = SeiLogin(chromedriver_path, chrome_options)
    prompt = PromptWindow(sei.root)
    
    def executar_selenium():
        prompt.prompt_window()
        time.sleep(4)
        
        
    selenium_thread = threading.Thread(target=executar_selenium)
    selenium_thread.start()
    sei.root.mainloop()