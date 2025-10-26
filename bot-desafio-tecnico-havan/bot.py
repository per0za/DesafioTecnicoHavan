import os
import logging
from pathlib import Path
from datetime import datetime

import pandas as pd

from webdriver_manager.firefox import GeckoDriverManager

from botcity.maestro import BotMaestroSDK
from botcity.web import WebBot, Browser, By

DIRETORIO_EXCEL = str(Path(__file__).resolve().parent / "ArquivoExcel")

class DownloadExcel:
    def __init__(self):
        self.bot = WebBot()
        self.maestro = BotMaestroSDK.from_sys_args()
        self._criar_log()

        self.df = pd.DataFrame()


    def main(self):
        logging.info("-------------------- Iniciando processo: Desafio técnico Havan --------------------")
        self.finalizar_instancia_navegador()

        self.configurar_bot()

        tentativas = 3
        for _ in range(tentativas, 0, -1):
            try:
                if not self.verificar_sessao_ativa():
                    self.configurar_bot()

                logging.info("Acessando a pagina do desafio tecnico...")
                self.bot.browse("https://www.rpachallenge.com/")
                logging.info("Pagina acessada!")

                logging.info("Validando a existencia do arquivo challenge.xlsx...")
                if os.path.exists(rf"{DIRETORIO_EXCEL}/challenge.xlsx"):
                    os.remove(rf"{DIRETORIO_EXCEL}/challenge.xlsx")
                    os.makedirs("ArquivoExcel", exist_ok=True)
                    logging.info(
                        "Arquivo challenge.xlsx encontrado, deletando o mesmo para baixar novamente..."
                    )

                botao_download_excel = self.bot.find_element(
                    "//a[contains(@href, 'challenge.xlsx') and contains(@class, 'uiColorPrimary') and contains(normalize-space(text()), 'Download Excel')]",
                    By.XPATH,
                )

                botao_download_excel.click()
                self.bot.wait(3000)
                logging.info("Download do arquivo concluido!")

                logging.info("Lendo o conteudo do arquivo challenge.xlsx...")
                self.df = pd.read_excel(rf"{DIRETORIO_EXCEL}/challenge.xlsx")

                self.df.columns = self.df.columns.str.strip().str.replace(
                    r"\s+", " ", regex=True
                )

                logging.info("Conteúdo lido! Dataframe criado com as informações do arquivo challenge.xlsx.")

                logging.info(
                    "Iniciando o processo de preenchimento do formulario do desafio..."
                )
                botao_start = self.bot.find_element(
                    "//button[contains(@class, 'uiColorButton') and normalize-space(text())='Start']",
                    By.XPATH,
                )
                botao_start.click()

                for index, row in self.df.iterrows():
                    logging.info("-----------------------------------------------------------------------------------")
                    logging.info(f"Iniciando as iterações de preenchimento de dados! Iteração nº {index + 1} ")
                    self.preencher_dados(row)

                logging.info("Preenchimento de todos os dados realizados com sucesso!")

                self.bot.wait(3000)
                self.bot.stop_browser()

                logging.info("------------------- Finalizando processo: Desafio técnico Havan -------------------")
                break

            except Exception as erro:
                logging.error(f"Erro ao tentar acessar a página: {erro}")
                logging.warning(f"Tentativas restantes: {tentativas - 1}")
                if tentativas == 0:
                    logging.error("Número máximo de tentativas foi atingido! Finalizando processo.")
                    break
                self.finalizar_instancia_navegador()
                self.configurar_bot()
                continue

    def preencher_dados(self, row):
        primeiro_nome = self.bot.find_element(
            "//input[@ng-reflect-name='labelFirstName']", By.XPATH
        )
        primeiro_nome.send_keys(row["First Name"])
        logging.info(f"Valor: {row["First Name"]}, informado no campo 'First Name'.")

        sobrenome = self.bot.find_element(
            "//input[@ng-reflect-name='labelLastName']", By.XPATH
        )
        sobrenome.send_keys(row["Last Name"])
        logging.info(f"Valor: {row["Last Name"]}, informado no campo 'Last Name'.")

        nome_empresa = self.bot.find_element(
            "//input[@ng-reflect-name='labelCompanyName']", By.XPATH
        )
        nome_empresa.send_keys(row["Company Name"])
        logging.info(f"Valor: {row["Company Name"]}, informado no campo 'Company Name'.")

        cargo = self.bot.find_element("//input[@ng-reflect-name='labelRole']", By.XPATH)
        cargo.send_keys(row["Role in Company"])
        logging.info(f"Valor: {row["Role in Company"]}, informado no campo 'Role in Company'.")

        endereco = self.bot.find_element(
            "//input[@ng-reflect-name='labelAddress']", By.XPATH
        )
        endereco.send_keys(row["Address"])
        logging.info(f"Valor: {row["Address"]}, informado no campo 'Address'.")

        email = self.bot.find_element(
            "//input[@ng-reflect-name='labelEmail']", By.XPATH
        )
        email.send_keys(row["Email"])
        logging.info(f"Valor: {row["Email"]}, informado no campo 'Email'.")

        numero_telefone = self.bot.find_element(
            "//input[@ng-reflect-name='labelPhone']", By.XPATH
        )
        numero_telefone.send_keys(row["Phone Number"])
        logging.info(f"Valor: {row["Phone Number"]}, informado no campo 'Phone Number'.")

        botao_enviar = self.bot.find_element(
            "//input[@type='submit' and contains(@class, 'uiColorButton') and @value='Submit']",
            By.XPATH,
        )

        botao_enviar.click()

    def configurar_bot(self):
        logging.info("Configurando bot...")
        self.bot = WebBot()
        self.bot.headless = True
        self.bot.browser = Browser.FIREFOX
        os.environ['WDM_LOG'] = str(logging.NOTSET)
        self.bot.driver_path = GeckoDriverManager().install()
        self.bot.download_folder_path = DIRETORIO_EXCEL
        self.bot.start_browser()
        logging.info("Bot configurado e navegador iniciado!")

    def finalizar_instancia_navegador(self):
        try:
            if self.bot:
                self.bot.stop_browser()
                logging.info("Instância do navegador finalizada com sucesso.")

        except Exception as erro:
            logging.warning(f"O navegador já está finalizado! Erro: {erro}")

    def verificar_sessao_ativa(self):
        logging.info("Verificando se a sessão do bot já está ativa...")
        try:
            _ = self.bot.driver.title
            logging.info("Verificação concluída, sessão já está ativa!")
            return True

        except Exception:
            logging.warning("Sessão perdida. Recriando bot...")
            return False

    def _criar_log(self):
        os.makedirs("logs", exist_ok=True)
        nome_arquivo_log = datetime.now().strftime("logs/execucao_%d-%m-%Y.txt")

        logging.basicConfig(
            filename=nome_arquivo_log,
            level=logging.INFO,
            format="%(asctime)s - %(levelname)s - %(message)s",
        )


if __name__ == "__main__":
    download_excel = DownloadExcel()
    download_excel.main()
