import subprocess  # Módulo para executar processos externos
import time  # Módulo para manipulação de tempo
import pyautogui  # Módulo para controle do mouse e teclado
import openpyxl  # Módulo para manipulação de arquivos Excel
import clipboard  # Módulo para interação com a área de transferência do sistema
from datetime import datetime  # Importa a classe datetime do módulo datetime

def abrir_roclink():
    """Abre o aplicativo Roclink."""
    # Caminho para o executável do Roclink
    caminho_roclink = r'"C:\Program Files (x86)\ROCLINK800\Roclink.exe"'
    # Abre o aplicativo Roclink com os parâmetros especificados
    subprocess.Popen(f'{caminho_roclink} /LOGIN:LOI:1000 /TCPIP:IP FROM DEVICE:4000 /MENU:view:History:From Device', shell=False)

def copiar_dados():
    """Realiza a operação de cópia dos dados e retorna o conteúdo copiado."""
    # Espera 10 segundos para que o aplicativo Roclink esteja pronto
    time.sleep(10)
    # Pressiona a tecla 'tab' três vezes
    pyautogui.press('tab', presses=3)
    # Pressiona a tecla 'down' para selecionar uma opção específica
    pyautogui.press('down')
    # If you wanna Daily Report uncomment the next line
    pyautogui.press('down')
    # Pressiona a tecla 'tab' para navegar para o próximo campo
    # pyautogui.press('tab')
    # Pressiona a tecla 'enter' para confirmar a seleção
    pyautogui.press('enter')
    # Aguarda 10 segundo para a cópia dos dados
    time.sleep(10)
    # Pressiona 'Ctrl + A' para selecionar todos os dados
    with pyautogui.hold('ctrl'):
        pyautogui.press(['a', 'c'])
    # Aguarda 10 segundos para garantir a cópia dos dados na área de transferência
    time.sleep(10)
    # Pressiona 'enter' para fechar a janela de visualização de dados
    pyautogui.press('enter')
    # Pressiona 'Alt + F4' para fechar a janela do aplicativo Roclink
    with pyautogui.hold('alt'):
        pyautogui.press('f4')
    # Retorna o conteúdo copiado da área de transferência
    return clipboard.paste()


def salvar_em_xlsx(dados):
    """Salva os dados copiados em um arquivo XLSX."""
    # Divide os dados em linhas
    linhas = dados.split('\n')
    # Obtém a data e hora atual formatada como string
    data_hora_atual = datetime.now().strftime("%d-%m-%Y_%H-%M-%S")
    # Nome do arquivo XLSX com a data e hora atual
    nome_arquivo_xlsx = f'Report-Roclink{data_hora_atual}.xlsx'
    # Cria um novo arquivo XLSX
    wb = openpyxl.Workbook()
    # Seleciona a planilha ativa no arquivo
    ws = wb.active
    # Especifica uma senha para que o arquivo gerado não possa ser alterado apênas consultado
    # ws.protection.password = 'Choose password if you want'
    # ws.protection.sheet = True
    # ws.protection.enable()

    # Para cada linha de dados
    for linha in linhas:
        # Divide os dados da linha em valores separados por tabulação
        dados_linha = linha.split('\t')
        # Remove o '\n' apenas da última coluna
        dados_linha[-1] = dados_linha[-1].rstrip('\n')
        # Converte os valores para float, se possível
        dados_linha = [float(valor) if valor.replace('.', '', 1).isdigit() else valor for valor in dados_linha]
        # Adiciona os valores como uma linha na planilha
        ws.append(dados_linha)

    # Salva o arquivo XLSX com o nome especificado
    wb.save(nome_arquivo_xlsx)
    # Imprime uma mensagem informando que os dados foram salvos
    print(f"Os dados foram salvos em '{nome_arquivo_xlsx}'.")


if __name__ == "__main__":
    try:
        # Tenta abrir o aplicativo, copiar os dados e salvá-los em um arquivo XLSX
        abrir_roclink()
        dados_copiados = copiar_dados()
        salvar_em_xlsx(dados_copiados)
    except Exception as e:
        # Se ocorrer um erro, imprime uma mensagem de erro
        print(f"Ocorreu um erro: {e}")
