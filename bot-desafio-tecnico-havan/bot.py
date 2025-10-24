import os
import logging
from pathlib import Path
from datetime import datetime

import pandas as pd

from botcity.maestro import BotMaestroSDK
from botcity.web import WebBot, Browser, By

DIRETORIO_RESOURCES = str(Path(__file__).resolve().parent / "resources")

class DownloadExcel:
    def __init__(self):
        self.bot = WebBot()
        self.maestro = BotMaestroSDK.from_sys_args()
        self._criar_log()

        self.df = pd.DataFrame()


    def main(self):
        self.bot.headless = False

        self.bot.browser = Browser.FIREFOX

        # TO DO
        self.bot.driver_path = self.bot.get_resource_abspath("geckodriver.exe")

        self.bot.download_folder_path = DIRETORIO_RESOURCES

        self.bot.browse("https://www.rpachallenge.com/")

        botao_download_excel = self.bot.find_element(
            "//a[contains(@href, 'challenge.xlsx') and contains(@class, 'uiColorPrimary') and contains(normalize-space(text()), 'Download Excel')]",
            By.XPATH,
        )
        botao_download_excel.click()
        self.bot.wait(3000)

        botao_start = self.bot.find_element(
            "//button[contains(@class, 'uiColorButton') and normalize-space(text())='Start']",
            By.XPATH,
        )
        botao_start.click()

        self.df = pd.read_excel(fr"{DIRETORIO_RESOURCES}/challenge.xlsx")

        self.df.columns = (
            self.df.columns
            .str.strip()
            .str.replace(r"\s+", " ", regex=True)
        )

        for _, row in self.df.iterrows():
            primeiro_nome = self.bot.find_element(
                "//input[@ng-reflect-name='labelFirstName']", By.XPATH
            )
            primeiro_nome.send_keys(row["First Name"])

            sobrenome = self.bot.find_element(
                "//input[@ng-reflect-name='labelLastName']", By.XPATH
            )
            sobrenome.send_keys(row["Last Name"])

            nome_empresa = self.bot.find_element(
                "//input[@ng-reflect-name='labelCompanyName']", By.XPATH
            )
            nome_empresa.send_keys(row["Company Name"])

            cargo = self.bot.find_element("//input[@ng-reflect-name='labelRole']", By.XPATH)
            cargo.send_keys(row["Role in Company"])

            endereco = self.bot.find_element(
                "//input[@ng-reflect-name='labelAddress']", By.XPATH
            )
            endereco.send_keys(row["Address"])

            email = self.bot.find_element("//input[@ng-reflect-name='labelEmail']", By.XPATH)
            email.send_keys(row["Email"])

            numero_telefone = self.bot.find_element(
                "//input[@ng-reflect-name='labelPhone']", By.XPATH
            )
            numero_telefone.send_keys(row["Phone Number"])

            botao_enviar = self.bot.find_element(
                "//input[@type='submit' and contains(@class, 'uiColorButton') and @value='Submit']",
                By.XPATH,
            )
            botao_enviar.click()

        self.bot.wait(3000)
        self.bot.stop_browser()


    def _criar_log(self):
        os.makedirs("logs", exist_ok=True)
        nome_arquivo_log = datetime.now().strftime("logs/execucao_%d-%m-%Y.log")

        logging.basicConfig(filename=nome_arquivo_log, level=logging.INFO, format="%(asctime)s - %(levelname)s - %(message)s")


if __name__ == "__main__":
    download_excel = DownloadExcel()
    download_excel.main()
