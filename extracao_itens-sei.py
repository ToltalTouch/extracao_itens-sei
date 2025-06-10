import logging
import os
import threading
import time
import pandas as pd
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import WebDriverException

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
download_dir = os.path.join(atual_dir, 'downloads')
if not os.path.exists(download_dir):
    os.makedirs(download_dir)

df = pd.read_excel(excel_path, sheet_name='Itens Extraídos').dropna(subset=['PROCESSO SEI'])
df = df[df['PROCESSO SEI'].str.strip() != '']
process_numbers = df['PROCESSO SEI'].tolist()

chrome_options = Options()

def encontrar_arquivos():
    sei.driver.get(sei_url)
    colunas = ["PROCESSO SEI", "NOME ARQUIVO", "NOME", "MATERIAL", "MODELO", "TAMANHO/GENERO", "QUANTIDADE"]
    
    if not os.path.exists(excel_path):
        df_cabecalho = pd.DataFrame(columns=colunas)
        df_cabecalho.to_excel(excel_path, index=False, sheet_name='Itens Extraídos')
        logging.info(f"Arquivo Excel criado em: {excel_path}")

    for process_number in process_numbers:
        processo_itens_extraidos = []
        try:
            WebDriverWait(sei.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'txtPesquisaRapida'))
            ).send_keys(process_number + Keys.RETURN)

            frame_lista = WebDriverWait(sei.driver, 10).until(
                lambda d: d.find_element(By.ID, "ifrArvore")
            )
            sei.driver.switch_to.frame(frame_lista)
            frm_arvore = WebDriverWait(sei.driver, 10).until(
                EC.presence_of_element_located((By.ID, 'frmArvore'))
            )
            itens_form = frm_arvore.find_elements(By.TAG_NAME, 'a')
            termos_encontrados = [item for item in itens_form if item.text.strip().startswith("Termo")]
            if termos_encontrados:
                nome_funcionario = ""
                logging.info(f"Processando termo: {len(termos_encontrados)} para o processo {process_number}")

                for idx, item_form in enumerate(termos_encontrados):
                    try:
                        documento_titulo = item_form.text.strip()
                        logging.info(f"Processando documento: {idx+1}/{len(termos_encontrados): {documento_titulo}}")

                        item_form.click()

                        frame_visualicao = WebDriverWait(sei.driver, 10).until(
                            lambda d: d.find_element(By.ID, "ifrVisualizacao")
                        )
                        sei.driver.switch_to.frame(frame_visualicao)

                        if not nome_funcionario:
                            try:
                                nome_funcionario_element = sei.driver.find_element(By.CLASS_NAME, "Texto_Justificado")
                                nome_funcionario = nome_funcionario_element.text
                                logging.info(f"Nome do funcionário encontrado: {nome_funcionario}")
                            except Exception as e:
                                logging.error(f"Erro ao localizar nome do funcionário: {e}")

                        try:
                            nome_funcionario_element = sei.driver.find_element(By.CLASS_NAME, "Texto_Justificado")
                            nome_funcionario = nome_funcionario_element.text
                            df.loc[df['PROCESSO SEI'] == process_number, 'NOME FUNCIONARIO'] = nome_funcionario
                        except Exception as e:
                            logging.error(f"Erro ao localizar nome do funcionário para o processo {process_number}: {e}")

                        try:
                            linhas_tabela = sei.driver.find_elements(By.XPATH, "//table//tr[position()>1]")
                            for linha in linhas_tabela:
                                celulas = linha.find_elements(By.TAG_NAME, 'td')
                                if len(celulas) >= 3 and celulas[1].text.strip():
                                    material = celulas[0].text.strip()
                                    modelo = celulas[1].text.strip()
                                    tamanho = celulas[2].text.strip()
                                    quantidade = ""

                                    if len(celulas) > 3:
                                        quantidade = celulas[3].text.strip()
                                    
                                    if modelo or tamanho or quantidade:
                                        itens_info ={
                                            "PROCESSO SEI": process_number,
                                            "NOME ARQUIVO": documento_titulo,
                                            "NOME": nome_funcionario,
                                            "MATERIAL": material,
                                            "MODELO": modelo,
                                            "TAMANHO/GENERO": tamanho,
                                            "QUANTIDADE": quantidade
                                        }
                                        processo_itens_extraidos.append(itens_info)
                        
                        except (Exception,WebDriverException) as e:
                            logging.error(f"Erro ao extrair itens do processo {process_number}: {e}")
                        sei.driver.switch_to.default_content()
                        sei.driver.switch_to.frame(frame_lista)
                    except (Exception, WebDriverException) as e:
                        logging.error(f"Erro ao processar o item {idx+1} do processo {process_number}: {e}")
                        sei.driver.switch_to.default_content()
                        sei.driver.switch_to.frame(frame_lista)
                        continue
                else:
                    logging.info(f"Nenhum termo encontrado para o processo {process_number}")
        except WebDriverException as e:
            logging.error(f"Erro ao processar o número do processo {process_number}: {e}")
            continue
        finally:
            if processo_itens_extraidos:
                try:
                    if os.path.exists(excel_path):
                        df_existente = pd.read_excel(excel_path, sheet_name='Itens Extraídos')
                        df_novos = pd.DataFrame(processo_itens_extraidos)
                        df_final = pd.concat([df_existente, df_novos], ignore_index=True)
                    else:
                        df_final = pd.DataFrame(processo_itens_extraidos)
                    
                    df_final.to_excel(excel_path, index=False, sheet_name='Itens Extraídos')
                    logging.info(f"Salvos {len(processo_itens_extraidos)} intes do processo {process_number}")
                
                except Exception as e:
                    backup_path = os.path.join(atual_dir, 'excel', f'backup_{process_number}_{int(time.time())}.xlsx')
                    pd.Dataframe(processo_itens_extraidos).to_excel(backup_path, index=False)
                    logging.error(f"Erro ao salvar os itens extraídos do processo {process_number} no Excel: {backup_path}")
            
            sei.driver.switch_to.default_content()
            time.sleep(1)
            
if __name__ == "__main__":
    sei = SeiLogin(chromedriver_path, chrome_options)
    prompt = PromptWindow(sei.root)
    def executar_selenium():
        prompt.prompt_window()
        sei.login(sei_url)
        encontrar_arquivos()
    selenium_thread = threading.Thread(target=executar_selenium)
    selenium_thread.start()