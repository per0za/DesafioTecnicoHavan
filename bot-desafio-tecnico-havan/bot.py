from pathlib import Path
from botcity.web import WebBot, Browser, By
from botcity.plugins.excel import BotExcelPlugin
from botcity.maestro import *

BotMaestroSDK.RAISE_NOT_CONNECTED = False

DIRETORIO_RESOURCES = str(Path(__file__).resolve().parent / "resources")

def main():
    maestro = BotMaestroSDK.from_sys_args()
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    bot.headless = False

    bot.browser = Browser.FIREFOX

    bot.driver_path = bot.get_resource_abspath("geckodriver.exe")

    bot.download_folder_path = DIRETORIO_RESOURCES

    bot.browse("https://www.rpachallenge.com/")

    botao_download_excel = bot.find_element("//a[contains(@href, 'challenge.xlsx') and contains(@class, 'uiColorPrimary') and contains(normalize-space(text()), 'Download Excel')]", By.XPATH)
    botao_download_excel.click()
    bot.wait(3000)

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
