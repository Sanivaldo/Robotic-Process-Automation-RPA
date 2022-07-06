from botcity.core import DesktopBot
import pyautogui
import time
import os
import pandas as pd
import signal
import psutil
from datetime import datetime
import shutil

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


        self.sleep(5000)

        # Abrido o aplicativo SAP

        self.execute(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

        self.wait(5000)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('1_Abriu_o_SAP.png')

        self.wait(2000)

        # Ir para o campo de pesquisa

        self.control_f(wait=3000)

        image = pyautogui.screenshot()
        im2 = pyautogui.screenshot('2_Caixa_de_pesquisa.png')

        # Pesquisar o CCP

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

        # ------- Colocar timer se aparecer o wsawoudblock

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

        # Inserir login e senha:

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

        # Selecionando as variaveis de extraçao

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

        if not self.find("clique_adiado", matching=0.97, waiting_time=10000):
            self.not_found("clique_adiado")
        self.click_relative(54, 16)

        if not self.find("encerrad_clique", matching=0.97, waiting_time=10000):
            self.not_found("encerrad_clique")
        self.click_relative(57, 13)

        if not self.find("clique_tipodenota", matching=0.97, waiting_time=10000):
            self.not_found("clique_tipodenota")
        self.click_relative(79, 12)

        pyautogui.press('TAB')

        #Voltar para VR

        self.paste('VR')

        if not self.find("clique_datadanota", matching=0.97, waiting_time=10000):
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

            if self.find("break_while", matching=0.97, waiting_time=10000):
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

        if not self.find("clique_tabela", matching=0.97, waiting_time=10000):
            self.not_found("clique_tabela")
        self.click_relative(9, 11)

        time.sleep(5)

        self.tab()

        if not self.find("clique_ok", matching=0.97, waiting_time=10000):
            self.not_found("clique_ok")
        self.click_relative(13, 13)

        time.sleep(1)

        if not self.find("clique_tabulacao", matching=0.97, waiting_time=10000):
            self.not_found("clique_tabulacao")
        self.click_relative(58, 15)

        time.sleep(1)

        if not self.find("clique_confirm", matching=0.97, waiting_time=10000):
            self.not_found("clique_confirm")
        self.click_relative(17, 17)

        time.sleep(1)

        if not self.find("click_confirm", matching=0.97, waiting_time=10000):
            self.not_found("click_confirm")
        self.click_relative(17, 15)

        while True:

            if self.find("find_excel", matching=0.97, waiting_time=10000):
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

        horario1 = datetime.now()

        print(horario1)

        hora1 = time.time

        print('Abriu o excel')

        data_nome = str(horario1.date())

        dados_nome = str('SE16N_IW59' + data_nome)

        self.paste(dados_nome)

        print('Colou o nome')

        time.sleep(3)

        self.enter()

        print('salvou em Documentos')

        time.sleep(90)

        fonte = (r'C:\Users\mis.lucas.coelho\Documents\SE16N_IW59' + data_nome + '.xlsx')

        print(fonte)

        time.sleep(90)

        print('Extração Encerrada')

        # --------------------------------------------------------------------------------------------------------------#

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

        #-------------------------------------------------------------------------------------------------------

        time.sleep(15)

        # Abrido o aplicativo do SAP
        self.execute(r"C:\Program Files (x86)\SAP\FrontEnd\SAPgui\saplogon.exe")

        time.sleep(5)

        # Ir para o campo de pesquisa
        pyautogui.hotkey('ctrl', 'f')

        time.sleep(5)

        # Pesquisar o CCP

        self.backspace()
        self.backspace()
        self.backspace()

        cod_1 = 'ccp'

        time.sleep(3)
        self.paste(cod_1)

        # Entrar no campo de login

        self.enter(wait=2)

        # -----------------------------#

        login = 'tr035129'
        senha = 'Neobpo@2026'

        # ------------------------------#

        # Inserir login e senha:

        time.sleep(5)

        if not self.find("Login", matching=0.97, waiting_time=10000):
            self.not_found("Login")
            self.click()

        time.sleep(3)

        self.paste(login)

        time.sleep(3)

        self.tab()

        time.sleep(3)

        self.paste(senha)

        time.sleep(3)

        self.enter(wait=1)

        # Selecionando as variaveis de extraçao

        time.sleep(5)

        cod_2 = ('se16n')
        self.paste(cod_2)
        self.enter()

        time.sleep(3)

        cod_3 = ('ZISTWM_NTDA01_WF')
        self.paste(cod_3)

        time.sleep(3)

        if not self.find("apagar_500", matching=0.97, waiting_time=10000):
            self.not_found("apagar_500")
        self.click_relative(21, 11)

        self.backspace()
        self.backspace()
        self.backspace()

        print("apagou o texto")

        self.enter(wait=5)

        if not self.find("click_qmnum", matching=0.97, waiting_time=10000):
            self.not_found("click_qmnum")
        self.click_relative(41, 12)

        self.shift_tab(wait=1)
        self.shift_tab(wait=1)

        self.enter(wait=5)

        ##Copiar a linha de Nota para o clipboard

        df = pd.read_excel(fonte, index_col=False)

        print(fonte)

        # Voltar time = 90seg
        time.sleep(90)

        df2 = df['Nota']

        df2.to_clipboard(header=None, index=False)

        # Voltar time = 90seg
        time.sleep(90)

        if not self.find("clique_no_criterios_selecao", matching=0.97, waiting_time=10000):
            self.not_found("clique_no_criterios_selecao")
        self.click_relative(137, 11)

        time.sleep(5)

        cont_1 = 9
        while True:
            if cont_1 == 0:
                break
            self.tab()
            cont_1 = cont_1 - 1

        self.enter(wait=5)

        print('colou as notas')

        # Voltar time = 30seg
        time.sleep(60)

        self.key_f8()

        # Voltar time = 120seg
        time.sleep(120)

        if not self.find("click_variante_exibicao", matching=0.97, waiting_time=10000):
            self.not_found("click_variante_exibicao")
        self.click_relative(76, 12)

        self.tab()

        cod_4 = "/NBPO_IW59"

        self.paste(cod_4)

        self.key_f8()

        while True:

            if self.find("notas_encontradas", matching=0.97, waiting_time=10000):
                print('Encontrou')
                time.sleep(5)
                break
            else:
                print('Esperando aparecer')
                time.sleep(2)

        print('Proxima etapa')

        # Voltar time = 120seg
        time.sleep(120)

        if not self.find("click_xlsx", matching=0.97, waiting_time=10000):
            self.not_found("click_xlsx")
        self.click_relative(32, 14)

        time.sleep(3)

        pyautogui.press('down')

        time.sleep(3)

        self.enter()

        time.sleep(2)

        if not self.find("click_avançar", matching=0.97, waiting_time=10000):
            self.not_found("click_avançar")
        self.click_relative(15, 15)

        time.sleep(3)

        while True:

            if self.find("clique_acesso_rapido", matching=0.97, waiting_time=10000):
                print('Encontrou')
                time.sleep(5)
                break
            else:
                print('Esperando aparecer')
                time.sleep(5)

        time.sleep(10)

        self.shift_tab(wait=3000)

        self.tab(wait=3300)

        horario1 = datetime.now()

        data_nome = str(horario1.date())

        self.paste('SE16N' + data_nome + ".MHTML")

        self.enter(wait=3000)

        print('Salvou em documentos')

        time.sleep(6)

        while True:

            if self.find("find_nota-operacao", matching=0.97, waiting_time=10000):
                print('Encontrou')
                time.sleep(5)
                break
            else:
                print('Esperando aparecer')
                time.sleep(3)

        # INSERIR ETAPA AQUI PARA SALVAR O XLSX NA PASTA DA REDE

        self.key_f12()

        self.sleep(3000)

        self.tab()

        self.sleep(3000)

        pyautogui.press('down')

        self.sleep(3000)

        if not self.find("click_pasta-excel", matching=0.97, waiting_time=10000):
            self.not_found("click_pasta-excel")
        self.click_relative(98, 5)

        self.sleep(3000)

        self.enter()

        self.sleep(3000)

        time.sleep(30)

        fonte = (r"C:\Users\mis.lucas.coelho\Documents\SAP\SAP GUI\SE16N" + data_nome + ".MHTML")

        time.sleep(2)

        print(fonte)

        destino = (
            r"\\10.221.243.11\rep_mis\RELATÓRIOS ENVIADOS\CLIENTE\COMGAS\BASES_RELATORIOS\BASES_SAP\Backup-SE16N")

        time.sleep(2)

        shutil.move(fonte, destino)

        time.sleep(2)

        fonte = (r"C:\Users\mis.lucas.coelho\Documents\SAP\SAP GUI\SE16N" + data_nome + ".xlsx")

        time.sleep(2)

        print(fonte)

        time.sleep(2)

        destino = (
            r"\\10.221.243.11\rep_mis\RELATÓRIOS ENVIADOS\CLIENTE\COMGAS\BASES_RELATORIOS\BASES_SAP\Backup-SE16N")

        time.sleep(2)

        os.system('taskkill /im EXCEL.exe')
        print('Fechou o Excel')

        shutil.move(fonte, destino)

        time.sleep(2)

        # Voltar time = 60seg
        time.sleep(60)

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

        time.sleep(3)

        print('Fim da extraçao')

        print('Fim da extraçao')

    def not_found(self, label):
        print(f"Element not found: {label}")


if __name__ == '__main__':
    Bot.main()


