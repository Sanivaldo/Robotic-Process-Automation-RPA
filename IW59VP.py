from botcity.core import DesktopBot
import pyautogui
import time
import shutil
import os
from datetime import datetime

# Uncomment the line below for integrations with BotMaestro
# Using the Maestro SDK
# from botcity.maestro import *


class Bot(DesktopBot):

    def action(self, execution=None):

        horario1 = datetime.now()

        print(horario1)

        hora1 = time.time()

        #Abrido o aplicativo do SAP
        self.execute(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

        time.sleep(3)

        #Ir para o campo de pesquisa

        self.control_f()

        #Pesquisar o CCP

        cod_1 = 'ccp'
        self.paste(cod_1)

        #Entrar no campo de login

        self.enter(wait=2)

        # -----------------------------#

        login = 'TR037956'
        senha = 'Neobpo@2023'

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
        cod_2 = 'iw59'
        self.paste(cod_2)
        self.enter()

        if not self.find( "clique_adiado", matching=0.97, waiting_time=10000):
            self.not_found("clique_adiado")
        self.click_relative(54, 16)

        if not self.find( "encerrad_clique", matching=0.97, waiting_time=10000):
            self.not_found("encerrad_clique")
        self.click_relative(57, 13)

        if not self.find( "clique_tipodenota", matching=0.97, waiting_time=10000):
            self.not_found("clique_tipodenota")
        self.click_relative(79, 12)

        pyautogui.press('TAB')

        self.paste('VP')

        if not self.find( "clique_datadanota", matching=0.97, waiting_time=10000):
            self.not_found("clique_datadanota")
        self.click_relative(72, 13)

        pyautogui.press('TAB')
        self.paste('01.01.2017')

       #Indo para o Final da Tela

        cont_1 = 54
        while True:
            if cont_1 == 0:
                break
            pyautogui.hotkey('shift', 'tab')
            cont_1 = cont_1 - 1
        self.paste('/NBPO_IW59')

        self.key_f8()

        while True:

            if self.find( "break_while", matching=0.97, waiting_time=10000):
                print('Encontrou')
                time.sleep(5)
                break
            else:
                print('Esperando aparecer')
                time.sleep(2)

        print('Proxima etapa')

        time.sleep(5)

        #--------------------------------------------#

        cont_2 = 26
        while True:
            if cont_2 == 0:
                break
            pyautogui.hotkey('tab')
            cont_2 = cont_2 - 1
            print(cont_2)

        if not self.find( "clique_tabela", matching=0.97, waiting_time=10000):
            self.not_found("clique_tabela")
        self.click_relative(9, 11)

        time.sleep(5)

        self.tab()

        if not self.find( "clique_ok", matching=0.97, waiting_time=10000):
            self.not_found("clique_ok")
        self.click_relative(13, 13)

        time.sleep(1)

        if not self.find( "clique_tabulacao", matching=0.97, waiting_time=10000):
            self.not_found("clique_tabulacao")
        self.click_relative(58, 15)

        time.sleep(1)

        if not self.find( "clique_confirm", matching=0.97, waiting_time=10000):
            self.not_found("clique_confirm")
        self.click_relative(17, 17)

        time.sleep(1)

        if not self.find( "click_confirm", matching=0.97, waiting_time=10000):
            self.not_found("click_confirm")
        self.click_relative(17, 15)
       
        while True:

            if self.find( "find_excel", matching=0.97, waiting_time=10000):
                print('Encontrou')
                time.sleep(5)
                break
            else:
                print('Esperando aparecer')
                time.sleep(2)

        print('Proxima etapa')

        time.sleep(5)

        self.key_f12()

        time.sleep(5)

        print('Abriu o excel')

        data_nome = str(horario1.date())

        self.paste('IW59_VP_' + data_nome)

        print('Colou o nome')

        time.sleep(3)

        if not self.find( "click_Documento", matching=0.97, waiting_time=10000):
            self.not_found("click_Documento")
        self.click_relative(85, 13)
        
        print('Mudou para Documentos')
        
        time.sleep(3)

        cont_3 = 11
        while True:
            if cont_3 == 0:
                break
            pyautogui.hotkey('tab')
            cont_3 = cont_3 - 1
            print(cont_3)

        self.enter()

        time.sleep(90)

        fonte = (r'C:\Users\mis.lucas.coelho\Documents\IW59_VP_' + data_nome + '.xlsx')

        print(fonte)

        destino = (r"\\10.221.243.11\rep_mis\RELATÓRIOS ENVIADOS\CLIENTE\COMGAS\BASES_RELATORIOS\BASES_SAP\Backup-IW59")

        shutil.move(fonte, destino)

        #os.remove(r"\\10.221.243.11\rep_mis\RELATÓRIOS ENVIADOS\CLIENTE\COMGAS\BASES_RELATORIOS\BASES_SAP\IW59.xlsx")

        print('Moveu o novo iw59_VP para nova pasta ')

        #fonte_final = (r"C:\Users\mis.lucas.coelho\Desktop\Automaçao\IW59.xlsx")

        #destino_final = (r"\\10.221.243.11\rep_mis\RELATÓRIOS ENVIADOS\CLIENTE\COMGAS\BASES_RELATORIOS\BASES_SAP")

        #shutil.move(fonte_final, destino_final)

        time.sleep(90)

        print('Extração Encerrada')

        hora2 = time.time()

        horario_duracao = hora1 - hora2

        print('O código durou' + str(horario_duracao))


        # Uncomment to mark this task as finished on BotMaestro
        # self.maestro.finish_task
        #     task_id=execution.task_id,
        #     status=AutomationTaskFinishStatus.SUCCESS,
        #     message="Task Finished OK."
        # )

    def not_found(self, label):
        print(f"Element not found: {label}")


class Fechar_programas(DesktopBot):
    def action(self, execution=None):

        os.system('taskkill /im EXCEL.exe')
        print('Fechou o Excel')
        time.sleep(2)

        os.system('taskkill /im saplogon.exe')
        print('Fechou o SAP')
        time.sleep(2)

        os.system('taskkill /im saplogon.exe')
        print('Fechou o SAP')
        time.sleep(2)

        os.system('taskkill /im saplogon.exe')
        print('Fechou o SAP')

        os.system('taskkill /im saplogon.exe')
        print('Fechou o SAP')


if __name__ == '__main__':
    Bot.main()

if __name__ == '__main__':
    Fechar_programas.main()









