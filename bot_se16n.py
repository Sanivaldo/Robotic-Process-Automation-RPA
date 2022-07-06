from botcity.core import DesktopBot
import pyautogui
import time
import shutil
from datetime import date
import os

# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(DesktopBot):

    def action(self, execution=None):

        # Abrido o aplicativo do SAP
        self.execute(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

        time.sleep(5)

        #Ir para o campo de pesquisa
        pyautogui.hotkey('ctrl','f')

        #Pesquisar o CCP

        cod_1 = 'ccp'
        self.paste(cod_1)

        #Entrar no campo de login

        self.enter(wait=2)

        # -----------------------------#

        login = 'tr035129'
        senha = 'Neobpo@2024'

        # ------------------------------#

        #Inserir login e senha:

        if not self.find( "Login", matching=0.97, waiting_time=10000):
            self.not_found("Login")
            self.click()


        self.paste(login)

        pyautogui.press('TAB')

        self.paste(senha)

        self.enter(wait=1)

        #Selecionando as variaveis de extraçao

        time.sleep(1)
        cod_2 = ('se16n')
        self.paste(cod_2)
        self.enter()

        time.sleep(5)

        cod_3 = ('ZISTWM_NTDA01_WF')
        self.paste(cod_3)












        # Uncomment to mark this task as finished on BotMaestro
        # self.maestro.finish_task

        #     task_id=execution.task_id,
        #     status=AutomationTaskFinishStatus.SUCCESS,
        #     message="Task Finished OK."
        # )

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()