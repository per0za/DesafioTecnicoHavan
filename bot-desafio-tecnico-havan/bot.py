from pathlib import Path
from botcity.web import WebBot, Browser, By
from botcity.plugins.excel import BotExcelPlugin
from botcity.maestro import *

BotMaestroSDK.RAISE_NOT_CONNECTED = False

DIRETORIO_RESOURCES = str(Path(__file__).resolve().parent / "resources")

def main():
    maestro = BotMaestroSDK.from_sys_args()
    execucao = maestro.get_execution()

    print(f"Task ID is: {execucao.task_id}")
    print(f"Task Parameters are: {execucao.parameters}")

    bot = WebBot()
    excel = BotExcelPlugin()

    bot.headless = False

    bot.browser = Browser.FIREFOX

    bot.driver_path = bot.get_resource_abspath("geckodriver.exe")

    bot.download_folder_path = DIRETORIO_RESOURCES

    bot.browse("https://www.rpachallenge.com/")

    botao_download_excel = bot.find_element("//a[contains(@href, 'challenge.xlsx') and contains(@class, 'uiColorPrimary') and contains(normalize-space(text()), 'Download Excel')]", By.XPATH)
    botao_download_excel.click()
    bot.wait(3000)

    botao_start = bot.find_element("//button[contains(@class, 'uiColorButton') and normalize-space(text())='Start']", By.XPATH)
    botao_start.click()

    dados_excel_desafio = excel.read(bot.get_resource_abspath("challenge.xlsx")).as_list()

    for dados in dados_excel_desafio[1:]:
        primeiro_nome = bot.find_element("//input[@ng-reflect-name='labelFirstName']", By.XPATH)
        primeiro_nome.send_keys(dados[0])

        sobrenome = bot.find_element("//input[@ng-reflect-name='labelLastName']", By.XPATH)
        sobrenome.send_keys(dados[1])

        nome_empresa = bot.find_element("//input[@ng-reflect-name='labelCompanyName']", By.XPATH)
        nome_empresa.send_keys(dados[2])

        cargo = bot.find_element("//input[@ng-reflect-name='labelRole']", By.XPATH)
        cargo.send_keys(dados[3])

        endereco = bot.find_element("//input[@ng-reflect-name='labelAddress']", By.XPATH)
        endereco.send_keys(dados[4])

        email = bot.find_element("//input[@ng-reflect-name='labelEmail']", By.XPATH)
        email.send_keys(dados[5])

        numero_telefone = bot.find_element("//input[@ng-reflect-name='labelPhone']", By.XPATH)
        numero_telefone.send_keys(dados[6])

        botao_enviar = bot.find_element("//input[@type='submit' and contains(@class, 'uiColorButton') and @value='Submit']", By.XPATH)
        botao_enviar.click()

    bot.wait(5000)

    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK.",
    #     total_items=0,
    #     processed_items=0,
    #     failed_items=0
    # )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
