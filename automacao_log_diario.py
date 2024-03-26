import stat
import shutil
from tkinter import *
from tkinter import messagebox
from tkinter import filedialog
from tkinter import font
from PIL import ImageTk, Image
from datetime import datetime
import time
import os
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.chrome.options import Options
import sys
import smtplib
from email.mime.text import MIMEText
from dotenv import load_dotenv
import ctypes
import subprocess

load_dotenv()  # take environment variables from .env.

script_directory = os.path.dirname(os.path.abspath(__file__))
# Configuração básica do log

erros = []

#URL
url = 'https://suporte.dominioatendimento.com/login.html'

#Senha para extração
senha_extracao = os.getenv("SENHA_EXTRACAO")

# Variáveis XPATH
xpath_input_usuario = '//*[@id="j_username"]'
xpath_input_senha = '//*[@id="j_password"]'
xpath_botao_login = '//*[@id="loginForm"]/div[3]/div/input'
xpath_botao_menu = '//*[@id="globalHeader"]/div[1]/button/i'
xpath_botao_backups = '//*[@id="sidebar"]/ul/li[6]/a'
xpath_botao_download = '//*[@id="DataTables_Table_0"]/tbody/tr/td[2]/span'
# Emails de destinatarios
destinatarios = [os.getenv("EMAIL_CLIENTE_1"), os.getenv("EMAIL_CLIENTE_2"), os.getenv("EMAIL_CLIENTE_3"), os.getenv("EMAIL_CLIENTE_4")]

def enviar_email(erros, destinatarios):
    # Configurações do servidor SMTP do Gmail e credenciais
    smtp_server = 'smtp.gmail.com'
    smtp_port = 587
    smtp_username = os.getenv("EMAIL_VENT")
    smtp_password = os.getenv("SENHA_EMAIL")
    # Configurações do e-mail
    de = os.getenv("EMAIL_VENT")
    assunto = '[ERROS AUTOMAÇÃO DIÁRIA] Ocorreu algum imprevisto na automação do backup diário de log'
    # Criar o objeto MIMEText
    erros = "\n • ".join(erros)
    mensagem = 'Olá! \nParece que a automação diária não foi executada com sucesso. Segue o log de erros:\n •' + erros + '\n ============== \nAtenciosamente, \nEquipe VENT.'
    mensagem_email = MIMEText(mensagem)
    # Configurar cabeçalhos do e-mail
    mensagem_email['From'] = de
    mensagem_email['Subject'] = assunto
    # Configurar a conexão SMTP
    with smtplib.SMTP(smtp_server, smtp_port) as servidor_smtp:
        # Iniciar a conexão com o servidor
        servidor_smtp.starttls()
        # Fazer login no servidor SMTP
        servidor_smtp.login(smtp_username, smtp_password)
        # Enviar o e-mail para cada destinatário na lista
        for para in destinatarios:
            mensagem_email['To'] = para
            servidor_smtp.sendmail(de, para, mensagem_email.as_string())
            print('E-mail enviado com sucesso!')

def limpar_pasta(caminho_pasta):
    # Verifica se a pasta existe antes de deletar
    try:
        if os.path.exists(caminho_pasta):
            shutil.rmtree(caminho_pasta)  # Deleta a pasta e seu conteúdo
            time.sleep(15)
        os.makedirs(caminho_pasta)  # Cria uma nova pasta limpa
        time.sleep(15)
    except Exception as e:
        erros.append(f"Falha ao limpar pasta: {caminho_pasta}")

def verifica_arquivo(caminho, nome_arquivo):
    try:
        existencia = False
        # Concatena o caminho com o nome do arquivo
        caminho_completo = os.path.join(caminho, nome_arquivo)
        # Verifica se o arquivo existe no caminho especificado
        if os.path.exists(caminho_completo):
            existencia = True
        return existencia
    except Exception as e:
        erros.append(f"Erro ao verificar arquivo {nome_arquivo}: {str(e)} -- {datetime.now()}")

def aplica_log(arquivo_db, arquivo_log):
    try:
        comando_aplica_log = f"dbeng16.exe {arquivo_db} -a {arquivo_log}"
        output_inicio = verifica_task('dbeng16')
        run_cmd_adm(comando_aplica_log, 300)
        output_depois = verifica_task('dbeng16')
        while output_inicio != output_depois:
            time.sleep(600)
            output_depois = verifica_task('dbeng16')
    except Exception as e:
        erros.append(f"Erro ao executar aplicação de log: {e}")

def verifica_task(task):
    try:
        # Iniciar o processo e capturar a saída
        process = subprocess.Popen(f'tasklist | find "{task}"', stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
        # Aguardar pela conclusão do processo
        output = process.communicate()
        # Imprimir a saída
        saida = output.decode('utf-8')
        return saida
    except Exception as e:
        erros.append(f"Erro ao executar o comando de verificar tasks: {e}")

def extrair_completo(pasta_download, destino_extracao, senha):
    try:
        if os.path.exists(pasta_download) and os.path.exists(destino_extracao):
            arquivos = os.listdir(pasta_download)
            if arquivos:
                arquivo_7z = os.path.join(pasta_download, arquivos[0])
                print(arquivo_7z)
                print("Extração completo em: ", datetime.now())
                run_cmd_adm(f"7z x {arquivo_7z} -p{senha} -o{destino_extracao}", 5000)
                print(datetime.now())
                print("Completo extraido")
            else:
                raise IndexError
        else:
            raise ValueError
    except IndexError as e:
        erros.append(f"Erro ao extrair o arquivo de backup completo, o diretório de extração está vazio ou o arquivo não foi baixado completamente: {str(e)} -- {datetime.now()}")
    except ValueError as e:
        erros.append(f"Erro ao extrair o arquivo de backup completo, o diretório de extração não existe ou o arquivo não foi baixado completamente: {str(e)} -- {datetime.now()}")
    except Exception as e:
        erros.append(f"Erro ao extrair o arquivo de backup completo: {str(e)}")

def extrair_log(arquivo_7z, destino_extracao, senha):
    try:
        run_cmd_adm(f"7z x {arquivo_7z} -p{senha} -o{destino_extracao}", 300)
        # Listar todos os arquivos no diretório
        lista_arquivos = os.listdir(destino_extracao)
        # Retornar a contagem de arquivos
        tamanho = len(lista_arquivos)
        tempo_espera = 120
        while tamanho < 5:
            time.sleep(tempo_espera)
            # Listar novamente todos os arquivos no diretório
            lista_arquivos = os.listdir(destino_extracao)
            # Retornar a contagem de arquivos novamente
            tamanho = len(lista_arquivos) 
            if tamanho < 5:
                if tempo_espera >= 60:
                    tempo_espera = tempo_espera//2
        print("Log extraido")
    except Exception as e:
        erros.append(f"Erro ao extrair o arquivo {arquivo_7z}: {str(e)} -- {datetime.now()}")

def run_cmd_adm(command, tempo_espera_comando):
    try:
        # ShellExecuteW é usado para executar um comando como administrador
        print("Comando a disparar: ", command)
        result = ctypes.windll.shell32.ShellExecuteW(None, "runas", "cmd.exe", f"/c {command}", None, 1)
        if result <= 32:
            raise Exception(f"Falha ao executar o comando. Código de erro: {result}")
        else:
            print("Comando executado com sucesso!")
            # Aguardar a conclusão do processo
            time.sleep(tempo_espera_comando)
    except Exception as e:
        erros.append(f"Falha ao iniciar o CMD automaticamente e disparar o comando: {command} erro:{str(e)} -- {datetime.now()}")

def centralizar_janela(janela, largura, altura):
    # Obter as dimensões da tela
    largura_tela = janela.winfo_screenwidth()
    altura_tela = janela.winfo_screenheight()

    # Calcular as coordenadas X e Y para centralizar a janela
    x = (largura_tela - largura) // 2
    y = (altura_tela - altura) // 2

    # Definir as dimensões e a posição da janela
    janela.geometry(f'{largura}x{altura}+{x}+{y}')

def escolher_destino_download():
    pasta_destino = filedialog.askdirectory(title="Escolha a pasta a qual o arquivo .dom será armazenado após o download")
    print(pasta_destino)
    pasta_destino = pasta_destino.replace("/", "\\")
    print(pasta_destino)
    return pasta_destino

def escolher_destino_extracao():
    pasta_destino = filedialog.askdirectory(title="Escolha a pasta a qual o arquivo .dom baixado será extraído")
    print(pasta_destino)
    return pasta_destino

def salvar_configuracoes(configuracao, valor):
    try:
        # Ler o conteúdo do arquivo de configurações
        with open('configuracoes.txt', 'r') as arquivo:
            linhas = arquivo.readlines()

        # Identificar a linha a ser modificada
        linha_alvo = None
        for i, linha in enumerate(linhas):
            if linha.startswith(f'{configuracao}='):
                linha_alvo = i
                break

        # Modificar a linha, se encontrada
        if linha_alvo is not None:
            nova_linha = f'{configuracao}={valor}\n'
            linhas[linha_alvo] = nova_linha

            # Escrever o novo conteúdo de volta no arquivo
            with open('configuracoes.txt', 'w') as arquivo:
                arquivo.writelines(linhas)
    except Exception as e:
        erros.append(f"Não foi possível salvar os dados fornecidos no arquivo configuracoes.txt: {str(e)} -- {datetime.now()}")

def criar_configuracoes():
    try:
        with open("configuracoes.txt", "w") as file:
            file.write('download_directory_log=\nextraction_directory_log=\ndownload_directory_completo=\nextraction_directory_completo=\n')
    except Exception as e:
        erros.append(f"Não foi possível criar o arquivo configuracoes.txt: {str(e)} -- {datetime.now()}")

def carregar_configuracoes():
    # Carrega os caminhos salvos do arquivo de configuração
    try:
        with open("configuracoes.txt", "r") as file:
            for line in file:
                key, value = line.strip().split("=")
                if key == "download_directory_log":
                    download_directory_log = value
                    if download_directory_log == '':
                        download_directory_log = None
                elif key == "extraction_directory_log":
                    extraction_directory_log = value
                    if extraction_directory_log == '':
                        extraction_directory_log = None
                elif key == "download_directory_completo":
                    download_directory_completo = value
                    if download_directory_completo == '':
                        download_directory_completo = None
                elif key == "extraction_directory_completo":
                    extraction_directory_completo = value
                    if extraction_directory_completo == '':
                        extraction_directory_completo = None
        return download_directory_log, extraction_directory_log, download_directory_completo, extraction_directory_completo
    except FileNotFoundError:
        return None, None, None, None
    except Exception as e:
        erros.append(f"Não foi possível carregar os dados contidos no arquivo configuracoes.txt: {str(e)} -- {datetime.now()}")

def click_escolha_download_log():
    valor = escolher_destino_download()
    salvar_configuracoes('download_directory_log', valor)
    if valor != '':
        botao_diretorio_download_log.config(text=valor)
    else:
        botao_diretorio_download_log.config(text='Escolha uma Pasta')

def click_escolha_extracao_log(): 
    valor = escolher_destino_extracao()
    salvar_configuracoes('extraction_directory_log', valor)
    if valor != '':
        botao_diretorio_extracao_log.config(text=valor)
    else:
        botao_diretorio_extracao_log.config(text='Escolha uma Pasta')

def click_escolha_download_completo():
    valor = escolher_destino_download()
    salvar_configuracoes('download_directory_completo', valor)
    if valor != '':
        botao_diretorio_download_completo.config(text=valor)
    else:
        botao_diretorio_download_completo.config(text='Escolha uma Pasta')

def click_escolha_extracao_completo(): 
    valor = escolher_destino_extracao()
    salvar_configuracoes('extraction_directory_completo', valor)
    if valor != '':
        botao_diretorio_extracao_completo.config(text=valor)
    else:
        botao_diretorio_extracao_completo.config(text='Escolha uma Pasta')

def click_iniciar():
    download_directory_log, extraction_directory_log, download_directory_completo, extraction_directory_completo = carregar_configuracoes()
    if download_directory_log == None or extraction_directory_log == None or download_directory_completo == None or extraction_directory_completo == None:
        messagebox.showinfo("Aviso", "Escolha as pastas de destino antes de realizar a automação. Esta configuração ficará salva.")
    else:
        if os.path.exists(download_directory_log) and os.path.exists(extraction_directory_log) and os.path.exists(download_directory_completo) and os.path.exists(extraction_directory_completo):
            janela.destroy()
            automacao_banco(download_directory_log, extraction_directory_log, download_directory_completo, extraction_directory_completo)
        else:
            messagebox.showinfo("Aviso", "Alguma pasta selecionada não existe mais")

def automacao_download(download_directory_log, destino_extracao_log):
    try:
        # Configuração das opções do Chrome
        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": download_directory_log,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        navegador = webdriver.Chrome(options=chrome_options)
        navegador.get(url)

        email = os.getenv("EMAIL")
        senha = os.getenv("SENHA")
        # Preenchendo os campos de usuário e senha
        username_element = navegador.find_element(By.XPATH, xpath_input_usuario)
        password_element = navegador.find_element(By.XPATH, xpath_input_senha)

        username_element.send_keys(email)
        password_element.send_keys(senha)

        # Enviando o formulário de login
        navegador.find_element(By.XPATH, xpath_botao_login).click()

        # Aguardando até que a página de login seja completamente carregada
        WebDriverWait(navegador, 10).until(
            EC.url_changes(url)
        )

        # Navegando para a página de backups
        navegador.find_element(By.XPATH, xpath_botao_menu).click()
        navegador.find_element(By.XPATH, xpath_botao_backups).click()

        # Filtrando backups de log
        navegador.find_element(By.XPATH, '//*[@id="menu-backup-bkp-realizado"]').click()

        navegador.find_element(By.XPATH, '//*[@id="backupForm"]/div[2]/div[3]/div/div/label[2]').click()

        navegador.find_element(By.XPATH, '//*[@id="backupForm:_idJsp0"]').click()
        time.sleep(3)
        # Coleta do nome do arquivo
        span_element = navegador.find_element(By.XPATH, '//*[@id="DataTables_Table_0"]/tbody/tr/td[2]/span')
        texto_span = span_element.get_attribute('textContent')
        # Coleta senha de extracao
        span_element_senha = navegador.find_element(By.XPATH, '//*[@id="DataTables_Table_0"]/tbody/tr[1]/td[4]')
        senha_extracao_log = span_element_senha.text
        # Caminho para o arquivo zip
        arquivo_zip = f'{download_directory_log}/{texto_span}'
        # Realizando download do arquivo zip completo (apenas do primeiro arquivo da tabela)
        navegador.find_element(By.XPATH, '//*[@id="DataTables_Table_0"]/tbody/tr/td[2]').click() #click para download
        baixado = False
        tentativa = 0
        tempo_estimado_segundos = 60*5
        #TESTE DE FINALIZAÇÃO DE DOWNLOAD
        while baixado == False:
            time.sleep(tempo_estimado_segundos) # Verifica download de acordo o tempo estimado
            baixado = verifica_arquivo(download_directory_log, texto_span)
            if baixado:
                break
            else:
                if tempo_estimado_segundos > 120:
                    tempo_estimado_segundos = tempo_estimado_segundos//2 # Caso não tenha baixado, testará novamente com o tempo dividido por 2
                tentativa += 1
            if tentativa == 45:
                raise TimeoutError
        navegador.quit()
        #Extrair arquivo baixado
        extrair_log(arquivo_zip, destino_extracao_log, senha_extracao_log)
    except TimeoutError as e:
        erros.append(f"Download do arquivo log não foi finalizado pois o tempo de espera foi muito longo, verifique a internet: {str(e)} -- {datetime.now()}")
    except Exception as e:
        erros.append(f"Download do arquivo log não foi finalizado, falha na automação web: {str(e)} -- {datetime.now()}")

def automacao_banco(download_directory_log, destino_extracao_log, download_directory_completo, destino_extracao_completo):
    # ..
    # ====== DEFINIÇÕES =======
    # ..
    # Caminhos para os dados utilizados
    nome_arquivo_log = 'contabil.log'
    nome_arquivo_db = 'contabil.db'
    caminho_pasta_dados = 'D:\\Dados'
    # Diretorios do banco de dados
    arquivo_log = os.path.join(caminho_pasta_dados, nome_arquivo_log)
    arquivo_db = os.path.join(caminho_pasta_dados, nome_arquivo_db)
    # Diretorio do Backup Diário
    arquivo_log_novo = os.path.join(destino_extracao_log, nome_arquivo_log)
    # Diretório do Backup Semanal
    arquivo_db_completo = os.path.join(destino_extracao_completo, nome_arquivo_db)
    arquivo_log_completo = os.path.join(destino_extracao_completo, nome_arquivo_log)
    # Comandos personalizados de cmd
    comando_stop = 'net stop "SQLANYs_Servidor_Dominio16"'
    comando_start = 'net start "SQLANYs_Servidor_Dominio16"'
    # Muda a permissão de leitura para escrita
    try:
        permissao_log = os.stat(arquivo_log).st_file_attributes
        permissao_db = os.stat(arquivo_db).st_file_attributes
        if permissao_log & (1 << 8):
            os.chmod(arquivo_log, os.stat(arquivo_log).st_mode | stat.S_IWRITE)
        if permissao_db & (1 << 8):
            os.chmod(arquivo_log, os.stat(arquivo_db).st_mode | stat.S_IWRITE)
    except Exception as e:
        erros.append(f"Falha ao tentar mudar a permissão dos arquivos contabil.db e contabil.log presentes na pasta Dados -- {datetime.now()}")
    # ..
    # ====== AÇÕES =======
    # ..
    # Extrai novamente o .dom do backup completo semanal armazenado
    extrair_completo(download_directory_completo, destino_extracao_completo, senha_extracao)
    # Comando STOP no banco de dados
    run_cmd_adm(comando_stop, 30)
    print("Stop no banco")
    #Substitui os dados da pasta download completo pela pasta dados
    run_cmd_adm(f"del /f {arquivo_db}", 15)
    run_cmd_adm(f"del /f {arquivo_log}", 15)
    run_cmd_adm(f"move /y {arquivo_db_completo} {caminho_pasta_dados}", 15)
    run_cmd_adm(f"move /y {arquivo_log_completo} {caminho_pasta_dados}", 15)
    # Aplica o log utilizando o comando fornecido
    print("Aplica log...")
    aplica_log(arquivo_db, arquivo_log)
    # Baixa e extrai log para pastas selecionadas
    print("Download log")
    automacao_download(download_directory_log, destino_extracao_log)
    # Substitui log da pasta Dados log pelo log baixado acima
    run_cmd_adm(f"del /f {arquivo_log}", 60)
    run_cmd_adm(f"move /y {arquivo_log_novo} {caminho_pasta_dados}", 60)
    # Aplica log novamente (demora mais)
    print("Aplica log...")
    aplica_log(arquivo_db, arquivo_log)
    # Inicia o banco de dados
    print("Iniciando banco...")
    run_cmd_adm(comando_start, 300)
    #Deleta a pasta de extração e cria uma limpa
    limpar_pasta(destino_extracao_log)
    if destino_extracao_log != (download_directory_log.replace("\\","/")):
        limpar_pasta(download_directory_log.replace("\\","/"))
    if len(erros) > 0:
        enviar_email(erros)
    quit()

def main():
    download_directory_log, extraction_directory_log, download_directory_completo, extraction_directory_completo = carregar_configuracoes()

    if download_directory_log == None or extraction_directory_log == None or download_directory_completo == None or extraction_directory_completo == None:
        if not os.path.exists('./configuracoes.txt'):
            criar_configuracoes()
        global janela 
        janela = Tk()
        janela.title("Automação Real Assessoria")

        # ícone da janela
        janela.iconbitmap('images\\vent_logo.ico')
            
        centralizar_janela(janela, 500, 400)

        janela.configure(bg="#FFFFFF")

        # Carregue a imagem usando o Pillow
        imagem = Image.open('images\\tela_imagem_diario.png')

        # Ajuste o tamanho da imagem para o tamanho desejado do rótulo
        largura_desejada = 500
        altura_desejada = 400
        imagem = imagem.resize((largura_desejada, altura_desejada))

        # Converta a imagem para um formato compatível com Tkinter
        imagem_tk = ImageTk.PhotoImage(imagem)

        # Crie um rótulo para exibir a imagem
        label_imagem = Label(janela, image=imagem_tk)
        # Ajuste as coordenadas e dimensões conforme necessário
        label_imagem.place(x=0, y=0, width=largura_desejada,
                        height=altura_desejada)
        
        global botao_diretorio_download_log
        global botao_diretorio_extracao_log
        global botao_diretorio_download_completo
        global botao_diretorio_extracao_completo
        botao_diretorio_download_log = Button(janela, text='Escolha uma Pasta',width=17, height=1, command=click_escolha_download_log)
        botao_diretorio_extracao_log = Button(janela, text='Escolha uma Pasta',width=17, height=1, command=click_escolha_extracao_log)
        botao_diretorio_download_completo = Button(janela, text='Escolha uma Pasta',width=17, height=1, command=click_escolha_download_completo)
        botao_diretorio_extracao_completo = Button(janela, text='Escolha uma Pasta',width=17, height=1, command=click_escolha_extracao_completo)
        botao_diretorio_download_log.place(x=90,y=185)
        botao_diretorio_extracao_log.place(x=90,y=275)
        botao_diretorio_download_completo.place(x=320,y=185)
        botao_diretorio_extracao_completo.place(x=320,y=275)
        botao_iniciar = Button(janela, 
                                command=click_iniciar, 
                                bg='#99ACFF',
                                highlightbackground='#99ACFF',
                                activebackground='#99ACFF',
                                width=15,  
                                height=1,  
                                relief=FLAT, 
                                font = font.Font(family="Open Sans", size=16),
                                text="Iniciar",
                                borderwidth=0, 
                                highlightthickness=0,
                                highlightcolor='#99ACFF',
                                fg="black")
        
        botao_iniciar.place(x=156,y=325)
        
        janela.mainloop()
    
    else:
        if os.path.exists(download_directory_log.replace("\\","/")) and os.path.exists(download_directory_completo.replace("\\","/")) and os.path.exists(extraction_directory_log) and os.path.exists(extraction_directory_completo):
            limpar_pasta(download_directory_log)
            automacao_banco(download_directory_log, extraction_directory_log, download_directory_completo, extraction_directory_completo)
        else:
            try:
                # Tenta excluir o arquivo
                os.remove('configuracoes.txt')
                main()
            except OSError as e:
                erros.append(f"Erro ao tentar excluir o arquivo configuracoes.txt que está corrompido ou com algum outro problema -- {datetime.now()}")
            except Exception as e:
                erros.append(f"Erro ao tentar excluir o arquivo configuracoes.txt que está corrompido ou com algum outro problema -- {datetime.now()}")

if __name__ == "__main__":
    main()
