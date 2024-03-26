import shutil
import stat
import subprocess
import sys
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
from dotenv import load_dotenv
import ctypes
import smtplib
from email.mime.text import MIMEText

load_dotenv()

script_directory = os.path.dirname(os.path.abspath(__file__))

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
    assunto = '[ERROS AUTOMAÇÃO] Ocorreu algum imprevisto na automação do backup completo semanal'
    # Criar o objeto MIMEText
    erros = "\n • ".join(erros)
    mensagem = 'Olá! \nParece que a automação semanal não foi executada com sucesso. Segue o log de erros:\n •' + erros + '\n ============== \nAtenciosamente, \nEquipe VENT.'
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

def verifica_arquivo(caminho, nome_arquivo):
    try:
        existencia = False
        # Concatena o caminho com o nome do arquivo
        caminho_completo = os.path.join(caminho, nome_arquivo)
        print('Verificando existencia de: ', caminho_completo)
        # Verifica se o arquivo existe no caminho especificado
        if os.path.exists(caminho_completo):
            existencia = True
        
        return existencia
    except Exception as e:
        erros.append(f"Não foi possível verificar existencia do arquivo {nome_arquivo} na pasta {caminho}: {e} -- Horário do ERRO: {datetime.now()}")

def extrair_completo(arquivo_7z, destino_extracao, senha):
    try:
        run_cmd_adm(f"7z x {arquivo_7z} -p{senha} -o{destino_extracao}", 7200)
        # verifica_validate = verifica_arquivo(destino_extracao, "validate.txt") 
        # while verifica_validate == False:
            # verifica_validate = verifica_arquivo(destino_extracao, "validate.txt") 
            # if verifica_validate == False:
            #     if tempo_espera >= 300:
            #         tempo_espera = tempo_espera//2
    except Exception as e:
        erros.append(f"A extração do arquivo ({arquivo_7z}) completo não ocorreu: {e} -- Horário do ERRO: {datetime.now()}")

#Função para chamar comandos via cmd
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

def criar_configuracoes():
    try:
        with open("configuracoes.txt", "w") as file:
            file.write('download_directory_log=\nextraction_directory_log=\ndownload_directory_completo=\nextraction_directory_completo=\n')
    except Exception as e:
        erros.append(f"Erro ao criar o arquivo configuracoes.txt -- ERRO: {e} -- Horário do ERRO: {datetime.now()}")

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
        erros.append(f"Erro ao salvar os dados no arquivo configuracoes.txt -- ERRO: {e} -- Horário do ERRO: {datetime.now()}")

def carregar_configuracoes():
    # Carrega os caminhos salvos do arquivo de configuração
    try:
        with open("configuracoes.txt", "r") as file:
            for line in file:
                key, value = line.strip().split("=")
                if key == "download_directory_completo":
                    download_directory = value
                    if download_directory == '':
                        download_directory = None
                elif key == "extraction_directory_completo":
                    extraction_directory = value
                    if extraction_directory == '':
                        extraction_directory = None
        return download_directory, extraction_directory
    except FileNotFoundError:
        return None, None
    except Exception as e:
        erros.append(f"Erro ao carregar os dados contidos no arquivo configuracoes.txt -- ERRO: {e} -- Horário do ERRO: {datetime.now()}")

def click_escolha_download():
    valor = escolher_destino_download()
    salvar_configuracoes('download_directory_completo', valor)
    if valor != '':
        botao_diretorio_download.config(text=valor)
    else:
        botao_diretorio_download.config(text='Escolha uma Pasta')

def click_escolha_extracao(): 
    valor = escolher_destino_extracao()
    salvar_configuracoes('extraction_directory_completo', valor)
    if valor != '':
        botao_diretorio_extracao.config(text=valor)
    else:
        botao_diretorio_extracao.config(text='Escolha uma Pasta')

def click_iniciar():
    download_directory, extraction_directory = carregar_configuracoes()
    if download_directory == None or extraction_directory == None:
        messagebox.showinfo("Aviso", "Escolha as pastas de destino antes de realizar a automação. Esta configuração ficará salva.")
    else:
        if os.path.exists(download_directory) and os.path.exists(extraction_directory):
            janela.destroy()
            automacao_download(download_directory, extraction_directory)
        else:
            messagebox.showinfo("Aviso", "Alguma pasta selecionada não existe mais")

def automacao_download(download_directory, destino_extracao):
    try:
        # Configuração das opções do Chrome
        chrome_options = Options()
        chrome_options.add_experimental_option("prefs", {
            "download.default_directory": download_directory,
            "download.prompt_for_download": False,
            "download.directory_upgrade": True,
            "safebrowsing.enabled": True
        })

        navegador = webdriver.Chrome(options=chrome_options)
        navegador.get(url)

        email = os.getenv('EMAIL')
        senha = os.getenv('SENHA')
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

        # Filtrando backups completos
        navegador.find_element(By.XPATH, '//*[@id="menu-backup-bkp-realizado"]').click()

        navegador.find_element(By.XPATH, '//*[@id="backupForm"]/div[2]/div[3]/div/div/label[3]').click()

        navegador.find_element(By.XPATH, '//*[@id="backupForm:_idJsp0"]').click()
        time.sleep(3)
        # Coleta do nome do arquivo
        span_element = navegador.find_element(By.XPATH, '//*[@id="DataTables_Table_0"]/tbody/tr/td[2]/span')
        texto_span = span_element.get_attribute('textContent')
        # Caminho para o arquivo zip
        arquivo_zip = f'{download_directory}\\{texto_span}'
        # Realizando download do arquivo zip completo (apenas do primeiro arquivo da tabela)
        try:
            # Tentar encontrar o elemento usando XPath
            elemento = navegador.find_element(By.XPATH, '//*[@id="DataTables_Table_0"]/tbody/tr/td[2]')
            # Verificar se o elemento existe
            while elemento == False:
                time.sleep(1800)
                navegador.refresh()
                elemento = navegador.find_element(By.XPATH, '//*[@id="DataTables_Table_0"]/tbody/tr/td[2]')
            navegador.find_element(By.XPATH, '//*[@id="DataTables_Table_0"]/tbody/tr/td[2]').click()
        except Exception as e:
            erros.append(f"Falha no download do arquivo, o arquivo completo não está disponível para download -- ERRO: {e} -- Horário do ERRO: {datetime.now()}")
            navegador.quit()
            quit()
        baixado = False
        tentativa = 0
        tempo_estimado_segundos = 60*60*6
        #TESTE DE FINALIZAÇÃO DE DOWNLOAD
        while baixado == False:
            time.sleep(tempo_estimado_segundos) # Verifica download de acordo com o calculo de tempo estimado
            baixado = verifica_arquivo(download_directory, texto_span)
            if baixado:
                break
            else:
                if tempo_estimado_segundos > 60*4:
                    tempo_estimado_segundos = tempo_estimado_segundos//2 # Caso não tenha baixado, testará novamente com o tempo
                tentativa += 1
            if tentativa == 60:
                raise TimeoutError
        navegador.quit()
        #Extrair_completo arquivo baixado
        extrair_completo(arquivo_zip, destino_extracao, senha_extracao)
        #Ao término da extração, realiza a próxima automação
        automacao_banco(destino_extracao)
        if len(erros) > 0:
            enviar_email(erros)
        quit()
    except TimeoutError as e:
        erros.append(f"Falha no download do arquivo de backup completo, ultrapassou o tempo limite de download, verifique sua conexão com a internet -- ERRO: {e} -- Horário do ERRO: {datetime.now()}")
    except Exception as e:
        erros.append(f"Falha no download do arquivo, falha na automação web -- ERRO: {e} -- Horário do ERRO: {datetime.now()}")

def automacao_banco(destino_extracao):
    # ..
    # ====== CONFIGURAÇÕES =======
    # ..
    # Caminhos para os dados utilizados no banco
    caminho_pasta_dados = 'D:\\Dados'
    arquivo_validate = os.path.join(destino_extracao, "validate.txt")
    caminho_dados_db = os.path.join(caminho_pasta_dados, "contabil.db")
    caminho_dados_log = os.path.join(caminho_pasta_dados, "contabil.log")
    # Comandos personalizados de cmd
    comando_stop = 'net stop "SQLANYs_Servidor_Dominio16"'
    comando_start = 'net start "SQLANYs_Servidor_Dominio16"'
    # ..
    # ====== AÇÕES PADRÕES =======
    # ..
    # Comando STOP no banco de dados
    run_cmd_adm(comando_stop, 60)
    if os.path.exists(arquivo_validate):
        os.remove(arquivo_validate)
    #Lista os arquivos faltantes dentro do diretório
    arquivos_alvo = os.listdir(destino_extracao)
    for arquivo in arquivos_alvo:
        arquivo_novo = os.path.join(destino_extracao, arquivo)
        arquivo_velho = os.path.join(caminho_pasta_dados, arquivo)
        # Muda a permissão de leitura para escrita
        permissao = os.stat(arquivo_velho).st_file_attributes
        if permissao & (1 << 8):
                os.chmod(arquivo_velho, os.stat(arquivo_velho).st_mode | stat.S_IWRITE)
    #Substitui os dados da pasta download completo pela pasta dados
        if os.path.exists(arquivo_velho) and os.path.exists(arquivo_novo):          
            run_cmd_adm(f"del /f {arquivo_velho}", 15)
            run_cmd_adm(f"move /y {arquivo_novo} {caminho_pasta_dados}", 15)
    aplica_log(caminho_dados_db, caminho_dados_log)
    #Inicia o banco de dados
    run_cmd_adm(comando_start, 300)

def main():
    download_directory, extraction_directory = carregar_configuracoes()

    if download_directory == None or extraction_directory == None:
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
        imagem = Image.open('images\\tela_imagem.png')

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
        
        global botao_diretorio_download
        global botao_diretorio_extracao
        botao_diretorio_download = Button(janela, text='Escolha uma Pasta',width=15, height=1, command=click_escolha_download)
        botao_diretorio_extracao = Button(janela, text='Escolha uma Pasta',width=15, height=1, command=click_escolha_extracao)
        botao_diretorio_download.place(x=100,y=248)
        botao_diretorio_extracao.place(x=330,y=248)

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
        if os.path.exists(download_directory.replace("\\","/")) and os.path.exists(extraction_directory):
            limpar_pasta(download_directory)
            automacao_download(download_directory, extraction_directory)
        else:
            try:
                # Tenta excluir o arquivo
                os.remove('configuracoes_completo.txt')
                # ou use os.unlink(caminho_do_arquivo)
                main()
            except OSError as e:
                erros.append(f"Falha ao deletar o arquivo configuracoes.txt -- ERRO: {e} -- Horário do ERRO: {datetime.now()}")
            except Exception as e:
                erros.append(f"Falha ao deletar o arquivo configuracoes.txt -- ERRO: {e} -- Horário do ERRO: {datetime.now()}")

if __name__ == "__main__":
    main()
