from botcity.core import DesktopBot
import pyautogui
import time
import shutil
import os
from datetime import datetime
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

        horario1 = datetime.now()

        print(horario1)

        hora1 = time.time

        self.sleep(5000)

        #Abrido o aplicativo SAP

        self.execute(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

        self.wait(5000)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('1_Abriu_o_SAP.png')

        self.wait(2000)

        #Ir para o campo de pesquisa

        self.control_f(wait=3000)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('2_Caixa_de_pesquisa.png')

        #Pesquisar o CCP

        self.wait(2000)

        cod_1 = 'ccp'

        self.backspace()
        self.backspace()
        self.backspace()

        self.paste(cod_1)

        self.wait(2000)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('3_Colou_o_ccp.png')

        self.wait(2000)

        #Entrar no campo de login
        
        if self.find("wsawoudblock", matching=0.97, waiting_time=10000):
            image = pyautogui.screenshot()
            im2 = pyautogui.screenshot('4_wsawoudblock.png')
            self.wait(2000)
            self.click_relative(181, 81)

        print("Continuou")
        
        self.enter(wait=5000)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('5_Tela_login.png')

        self.wait(2000)

        login = 'TR037956'
        senha = 'Neobpo@2023'

        #Inserir login e senha:

        if not self.find("Login", matching=0.97, waiting_time=10000):
            self.not_found("Login")
        self.click()

        self.wait(2000)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('6_Tela_login_Achar_campos_de_login_senha.png')

        self.wait(2000)

        self.paste(login)

        pyautogui.press('TAB')

        self.paste(senha)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('7_Tela_login_Achar_campos_de_login_senha_inseridos.png')

        self.enter(wait=4000)

        #Selecionando as variaveis de extraçao

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('8_Tela_usuarios.png')

        time.sleep(3)

        cod_2 = 'iw59'
        self.paste(cod_2)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('8_Tela_usuarios_colou_iw59.png')

        self.enter()

        self.wait(2000)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('8_Tela_usuarios_selecao_de_notas.png')

        self.wait(2000)

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

        self.paste('VR')

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
        hora_nome = str(horario1.hour)
        dados_nome = str('IW59_' + data_nome + "-" + hora_nome + "h")
        dados_fonte = (data_nome + "-" + hora_nome + "h")

        self.paste(dados_nome)

        print('Colou o nome')

        time.sleep(3)

        self.enter()
        
        print('salvou em Documentos')

        time.sleep(90)

        fonte = (r'C:\Users\mis.lucas.coelho\Documents\IW59_' + dados_fonte + '.xlsx')

        print(fonte)

        destino = (r"\\10.221.243.11\rep_mis\RELATÓRIOS ENVIADOS\CLIENTE\COMGAS\BASES_RELATORIOS\BASES_SAP\Backup-IW59")

        shutil.move(fonte, destino)

        print('Moveu o novo iw59 para nova pasta ')

        time.sleep(90)

        print('Extração Encerrada')

        #--------------------------------------------------------------------------------------------------------------#

        os.system('taskkill /im EXCEL.exe')
        print('Fechou o Excel')

        time.sleep(90)

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

        time.sleep(5)

        # Ir para o campo de pesquisa

        self.control_f()

        # Pesquisar o CCP

        cod_1 = 'ccp'
        self.paste(cod_1)

        # Entrar no campo de login

        self.enter(wait=2)

        # -----------------------------#

        login = 'TR037956'
        senha = 'Neobpo@2023'

        # ------------------------------#

        # Inserir login e senha:

        if not self.find( "Login", matching=0.97, waiting_time=10000):
            self.not_found("Login")
        self.click()

        self.paste(login)

        pyautogui.press('TAB')

        self.paste(senha)

        self.enter(wait=1)

        # Selecionando as variaveis de extraçao

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

        # Indo para o Final da Tela

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

        # --------------------------------------------#

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

        hora_nome = str(horario1.hour)

        dados_nome = str('IW59_VP_' + data_nome + "-" + hora_nome + "h")

        dados_fonte = (data_nome + "-" + hora_nome + "h")

        self.paste(dados_nome)

        print('Colou o nome')

        time.sleep(3)

        self.enter()

        print('Salvou em Documentos')

        time.sleep(90)

        fonte = (r'C:\Users\mis.lucas.coelho\Documents\IW59_VP_' + dados_fonte + '.xlsx')

        print(fonte)

        destino = (r"\\10.221.243.11\rep_mis\RELATÓRIOS ENVIADOS\CLIENTE\COMGAS\BASES_RELATORIOS\BASES_SAP\Backup-IW59")

        shutil.move(fonte, destino)

        print('Moveu o novo iw59_VP para nova pasta ')

        time.sleep(10)

        print('Extração Encerrada')

        os.system('taskkill /im EXCEL.exe')
        print('Fechou o Excel')

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


















