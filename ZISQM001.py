from botcity.core import DesktopBot
import pyautogui
import time
from datetime import datetime, date
import shutil
import os
import psutil
import signal

class Bot(DesktopBot):

    def action(self, execution=None):

        Programa = r'saplogon.exe'

        def findProcess(name):

            procs = list()
            for proc in psutil.process_iter():
                try:
                    if proc.name() == name and proc.status() == psutil.STATUS_RUNNING:
                        pid = proc.pid
                        procs.append(pid)
                        print(pid)
                except:
                    pass
            return procs

        processo = findProcess(Programa)

        for i in processo:
            os.kill(i, signal.SIGTERM)

        # Abrido o aplicativo do SAP

        self.execute(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

        time.sleep(3)

        #Ir para o campo de pesquisa
        pyautogui.hotkey('ctrl','f')

        #Pesquisar o Ccp

        self.key_esc()
        self.backspace()
        self.backspace()
        self.backspace()

        cod_1 = 'ccp'
        self.paste(cod_1)

        #Entrar no campo de login

        self.enter(wait=2)

        # -----------------------------#

        login = 'TR037956'
        senha = 'Neobpo@2023'

        # ------------------------------#

        #Inserir login e senha:

        if not self.find("Login", matching=0.97, waiting_time=10000):
            self.not_found("Login")
            self.click()

        self.paste(login)

        pyautogui.press('TAB')

        self.paste(senha)

        self.enter(wait=1)

        #Selecionando as variaveis de extraçao

        time.sleep(2)

        cod_2 = ('ZISQM001')
        self.paste(cod_2)
        self.enter()

        time.sleep(2)

        if not self.find("click_status_da_mobilidade", matching=0.97, waiting_time=10000):
            self.not_found("click_status_da_mobilidade")
        self.click_relative(131, 15)
        
        if not self.find("click_nota", matching=0.97, waiting_time=10000):
            self.not_found("click_nota")
        self.click_relative(37, 11)
        
        self.tab()

        self.paste('*7750*')

        if not self.find("click_data_criacao", matching=0.97, waiting_time=10000):
            self.not_found("click_data_criacao")
        self.click_relative(88, 12)
        
        self.tab()
            
        self.paste('01.01.2019')

        self.tab()

        data = date.today()

        data_paste = (str(data.day) + '.' + str(data.month) + '.' + str(data.year))

        self.paste(str(data_paste))

        self.key_f8()

        while True:

            if self.find("find_tela_dados", matching=0.97, waiting_time=10000):
                print('Encontrou')
                time.sleep(5)
                break
            else:
                print('Esperando tela dados')
                time.sleep(2)

        self.key_f9()

        while True:

            if self.find("find_relatorio_de_controle", matching=0.97, waiting_time=10000):
                print('Encontrou')
                time.sleep(5)
                break
            else:
                print('Esperando aparecer')
                time.sleep(2)

        horario1 = datetime.now()

        data_nome = str(horario1.date())

        self.paste('ZISQM001_' + data_nome + '.txt')

        print('Colou o nome')

        self.shift_tab(wait=1)

        folder = (r'C:\Users\mis.lucas.coelho\Desktop\Automaçao')

        self.paste(folder,wait=1)

        if not self.find("click_gerar", matching=0.97, waiting_time=10000):
            self.not_found("click_gerar")
        self.click_relative(38, 16)

        #time = 300seg
        time.sleep(300)

        try:

            fonte= (r'C:\\Users\mis.lucas.coelho\Desktop\Automaçao\ZISQM001_' + data_nome +'.txt')

        except:

            print('Erro de fonte')

        destino = (r'\\10.221.243.11\rep_mis\RELATÓRIOS ENVIADOS\CLIENTE\COMGAS\BASES_RELATORIOS\BASES_SAP\Backup-ZISQM001')

        shutil.move(fonte, destino)

        time.sleep(10)

        os.system('taskkill /im EXCEL.exe')
        print('Fechou o Excel')

        os.system('taskkill /im EXCEL.exe')
        print('Fechou o Excel')

        time.sleep(10)

        Programa = r'saplogon.exe'

        def findProcess(name):

            procs = list()
            for proc in psutil.process_iter():
                try:
                    if proc.name() == name and proc.status() == psutil.STATUS_RUNNING:
                        pid = proc.pid
                        procs.append(pid)
                        print(pid)
                except:
                    pass
            return procs

        processo = findProcess(Programa)

        for i in processo:
            os.kill(i, signal.SIGTERM)


    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()








































































































































