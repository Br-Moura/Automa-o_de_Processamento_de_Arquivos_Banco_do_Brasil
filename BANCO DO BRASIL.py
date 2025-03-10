import re
import os
import shutil
import pyautogui
import sys
import time
import subprocess
from tkinter import Tk, messagebox
import tkinter as tk
from tkinter.filedialog import askopenfilename

# Dados fixos para o produto: Banco Brasil
STR_LAYOUT = 'C:\MATICA\MACHINE\MP_Layouts\BANCO_BRASIL.mpl'
STR_XML = 'C:\MATICA\MACHINE\MP_Databases\XML\BANCO_BRASIL\BB_ALTUS.XML'
STR_BD = 'C:\MATICA\MACHINE\MP_Databases\PRD\BANCO_BRASIL'
STR_LETTER = ''
STR_TITLE = ' ** AVISO ** ==> SELECIONE A DATA DO MOVIMENTO BANCO BRASIL'

# Caminho para o repositório do AutoDec
AUTODEC_PATH = 'C:\AutoDec\CHDE_AutoEncDec.exe'
AUTODEC_INPUT_FOLDER = 'C:\AutoDec\entrada'
AUTODEC_OUTPUT_FOLDER = 'C:\AutoDec\saida'

# Dados fixos
root = tk.Tk()
root.withdraw()

# Função para abrir diálogo de seleção de arquivo
def select_file():
    Tk().withdraw()
    initialdir_str = 'U:\PAYMENT\BDB\VS\MUL_DI_CJ811_METAL\Matica'
    title_str = STR_TITLE
    file_fullname = askopenfilename( initialdir = initialdir_str, title = title_str, \
                                 filetypes = (('TXT files', '*.txt'), ('all files','*.*')) )
    print(file_fullname)
    if os.path.isfile(file_fullname):
        print(f'Arquivo Matica selecionado: {file_fullname}')
    else:
        print(f'Arquivo selecionado não encontrado: {file_fullname}')
        exit()
    return file_fullname

# Encontrar arquivo PRD correspondente
def find_prd_file(file_fullname):
    file_name = os.path.basename(file_fullname)
    dir_name = os.path.dirname(file_fullname)

    if file_name.endswith('.txt'):
        file_name = file_name[:-4]

    pos_date = dir_name.rfind('/')
    MC_dir = 'U:\PAYMENT\BDB\VS\MUL_DI_CJ811_METAL' + dir_name[pos_date:]
    prdfile_name = os.path.join(MC_dir, file_name + '.PRD.ENC')

    print(f'Arquivo PRD.ENC esperado: {prdfile_name}')

    if os.path.isfile(prdfile_name):
        print(f'ARQUIVO PRD.ENC EXISTE: {prdfile_name}')
        print(f'Arquivo: {file_name}')
    else:
        msg_str = f'Arquivo PRD.ENC não existe: {prdfile_name}'
        print(msg_str)
        messagebox.showerror('** PROBLEMA **', msg_str)
        exit()

    return prdfile_name

# Função para copiar arquivo para um diretório de input do AutoDec
def copy_file_to_autodec(prd_enc_file):
    try:
        local_file_name = shutil.copy(prd_enc_file, AUTODEC_INPUT_FOLDER)
        print(f'Arquivo copiado para a pasta de input do AutoDec: {local_file_name}')
        return local_file_name
    except Exception as e:
        msg_str = 'Erro ao copiar arquivo para autodec'
        messagebox.showerror('ERRO', msg_str)
        sys.exit()

# Função para descriptografar o arquivo usando o AutoDec
def decrypt_file(prd_enc_file):
    input_file = os.path.join(AUTODEC_INPUT_FOLDER, os.path.basename(prd_enc_file))
    output_folder = AUTODEC_OUTPUT_FOLDER

    # Copie o arquivo para a pasta de input do AutoDec
    shutil.copy(prd_enc_file, input_file)
    print(f'Arquivo PRD.ENC copiado para a pasta de input do AutoDec: {input_file}')

    # Comando para executar o AutoDec
    command = [
        AUTODEC_PATH,
        input_file,
        AUTODEC_INPUT_FOLDER,
        output_folder
    ]

    # Execute o comando
    process = subprocess.Popen(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
    stdout, stderr = process.communicate()

    try:
        stdout = stdout.decode('utf-8')
    except UnicodeDecodeError:
        stdout = stdout.decode('latin-1')
    try:
        stderr = stderr.decode('utf-8')
    except UnicodeDecodeError:
        stderr = stderr.decode('latin-1')

    if process.returncode == 0:
        print('Arquivo descriptografado com sucesso.')
        print('Saída do comando:', stdout)
    else:
        print('Erro ao descriptografar o arquivo.')
        print('Saída do comando:', stdout)
        print('Erro do comando:', stderr)
        raise subprocess.CalledProcessError(process.returncode, command)
    time.sleep(2.0)


# Função para encontrar o arquivo descriptografado na pasta de output
def find_decrypted_file(original_file):
    print("Carregando...")
    time.sleep(2.5)
    base_name = os.path.splitext(os.path.basename(original_file))[0]  # Obtém o nome base do arquivo original
    decrypted_file_name = base_name   # Nome esperado do arquivo descriptografado
    decrypted_file_path = os.path.join(AUTODEC_OUTPUT_FOLDER, decrypted_file_name)  # Caminho completo do arquivo descriptografado
    
    if os.path.isfile(decrypted_file_path):
        print(f'Arquivo descriptografado encontrado: {decrypted_file_path}')
        return decrypted_file_path
    else:
        raise FileNotFoundError(f'Nenhum arquivo descriptografado encontrado com o nome: {decrypted_file_name}')

# Função para analisar linha do arquivo Matica e extrair o número do cartão
def parse_matica_line(line):
    try:
        # Ajustar regex para capturar número de registro antes de #EMB01# e número do cartão após #EMB05#
        match = re.search(r'(\d{7})\$.*#EMB1#(\d{4} \d{4} \d{4} \d{4})', line)
        if match:
            register_number = match.group(1)
            card_number = match.group(2).replace(' ', '')  
            return card_number, register_number
        else:
            return None, None
    except Exception as e:
        msg_str = 'Falha ao analisar registro Mática, existem registros fora do padrão.'
        print(f"Erro: {e}")  
        sys.exit()

def parse_prd_record(record):
    try:
        # Ajustar regex para capturar número de registro antes de #EMB01# e número do cartão após #EMB05#
        match = re.search(r'(\d{7})\.*#EMB1#(\d{4} \d{4} \d{4} \d{4})', record)
        if match:
            register_number = match.group(1)
            card_number = match.group(2).replace(' ', '')  # Remove espaços no número do cartão
            return card_number, register_number
        else:
            return None, None
    except Exception as e:
        msg_str = 'Falha ao analisar registro PRD, PAN fora do padrão.'
        print(f"Erro: {e}")  
        sys.exit()


# Função para comparar os arquivos Matica e PRD
def compare_files(matica_file, prd_file):
    print("Iniciando a comparação dos arquivos...")

    encodings = ['utf-8', 'ISO-8859-1', 'cp1252']
    matica_records = []

    # Ler o arquivo Matica com diferentes codificações
    for encoding in encodings:
        try:
            with open(matica_file, 'r', encoding=encoding) as m_file:
                for line in m_file:
                    card_number, register_number = parse_matica_line(line)
                    if card_number and register_number:
                        matica_records.append((card_number, register_number))
                print(f'Arquivo Matica lido com codificação {encoding}')
                break
        except UnicodeDecodeError:
            print(f'Falha ao abrir o arquivo Matica com codificação {encoding}')
        except Exception as e:
            print(f'Falha ao abrir o arquivo Matica: {e}')
            sys.exit()

    prd_records = set()
    encodings = ['utf-8', 'ISO-8859-1']

    # Ler o arquivo PRD com diferentes codificações
    for encoding in encodings:
        try:
            with open(prd_file, 'r', encoding=encoding) as p_file:
                for record in p_file:
                    card_number, register_number = parse_prd_record(record)
                    if card_number and register_number:
                        prd_records.add((card_number, register_number))
                print(f'Número de registros no arquivo PRD com codificação {encoding}: {len(prd_records)}')
                if len(prd_records) > 0:
                    break
        except UnicodeDecodeError:
            print(f'Falha ao abrir o arquivo PRD com codificação {encoding}')
        except Exception as e:
            print(f'Falha ao abrir o arquivo PRD: {e}')
            sys.exit()

    if not matica_records:
        msg_str = 'Não foi possível abrir o arquivo Matica com nenhuma das codificações esperadas'
        print(msg_str)
        messagebox.showerror('ERRO', msg_str)
        sys.exit()

    matica_set = set(matica_records)
    print(f'Número total de registros extraídos do Matica: {len(matica_set)}')

    if not prd_records:
        msg_str = 'Não foi possível ler registros válidos do arquivo PRD'
        print(msg_str)
        messagebox.showerror('ERRO', msg_str)
        sys.exit()

    # Comparação dos registros
    for matica_record in matica_records:
        if matica_record in prd_records:
            print(f'Registros correspondentes encontrados')
        else:
            msg_str = f'Registro não encontrado no PRD'
            print(msg_str)
            messagebox.showerror('** ERRO **', msg_str)
            sys.exit()

    # Checar se todos os registros Matica foram encontrados no PRD
    if len(matica_set) == len(prd_records):
        msg_str = 'Todos os registros do Matica foram encontrados no PRD. Comparação bem-sucedida!'
        print(msg_str)
        
    else:
        msg_str = 'Discrepância encontrada: nem todos os registros Matica têm correspondentes no PRD.'
        print(msg_str)
        messagebox.showerror('** ERRO **', msg_str)
        sys.exit()

    print(f'Total de correspondências encontradas: {len(matica_set & prd_records)}')



def copy_files_for_matica(file_fullname):
    shutil.copy(file_fullname, STR_BD)
# Função para deletar arquivo descriptografado
def delete_file(file):
    os.remove(file)
    print(f'Arquivo excluído: {file}')

# Função main para chamar os métodos do script
def automate_maticard(file_write, local_fullfile_name):
    pyautogui.press('winleft')
    pyautogui.write('Maticard pro')
    pyautogui.press('enter')
    time.sleep(1.3)

    with pyautogui.hold('winleft'):
        pyautogui.press('left')     # Maticard Pro no canto esquerdo
    time.sleep(1.3)

    pyautogui.click(50, 83)         # Maticard Pro - Job Edit
    time.sleep(1.3)

    pyautogui.press(['tab','tab','tab','tab','tab','tab'])
    pyautogui.write(file_write)
    pyautogui.press('tab')
    pyautogui.write(file_write)

    # 5 - Selecionar BD, XML e importar 
    with pyautogui.hold('ctrl'):
        pyautogui.press('tab')

    pyautogui.press('enter')

    #pyautogui.click(677, 165)
    #time.sleep(3)
    pyautogui.click(579, 234)
    time.sleep(0.3)
    pyautogui.click(579, 310)
    time.sleep(0.3)
    pyautogui.click(579, 370)
    time.sleep(0.3)
    #pyautogui.press('down','down','down','down')
    #pyautogui.press('enter')

        
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.write(f'{STR_BD}\{file_write}.TXT')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('enter')
    time.sleep(1)
    pyautogui.press('tab')
    time.sleep(1)
    pyautogui.write(STR_XML)
    pyautogui.press(['tab','tab','tab'])
    pyautogui.press('enter')

    time.sleep(1.5)
    pyautogui.press('enter')

    #pyautogui.press(['tab','tab'])
    #pyautogui.press('downKey')D039C1GX00101308600
    #pyautogui.click(1096, 569)
    pyautogui.click(1096, 705)
    time.sleep(0.3)
    pyautogui.press('down')
    time.sleep(0.3)
    pyautogui.press('enter')
    time.sleep(0.3)
    pyautogui.press(['tab','tab'])
    pyautogui.press('enter')
    time.sleep(1.0)
    pyautogui.press('tab')
    time.sleep(1.0)
    pyautogui.write(STR_LAYOUT)

    pyautogui.press(['tab','tab'])
    pyautogui.press('enter')
    time.sleep(1.5)

    pyautogui.click(588, 275)
    time.sleep(0.3)

    pyautogui.doubleClick(1294, 433)
    time.sleep(0.3)
    pyautogui.press('tab')
    pyautogui.press('enter')
    pyautogui.write(file_write)
    #pyautogui.write(STR_LETTER)
    pyautogui.press(['tab','tab'])
    pyautogui.press('enter')
    time.sleep(0.3)
    pyautogui.press('tab')
    pyautogui.press('enter') 

# Função principal
def main():
    file_fullname = select_file()
    prd_enc_file = find_prd_file(file_fullname)
    copy_file_to_autodec(prd_enc_file)
    decrypt_file(prd_enc_file)
    decrypted_file = find_decrypted_file(prd_enc_file)

    file_write = os.path.splitext(os.path.basename(file_fullname))[0]
    print(f'Nome do arquivo a ser escrito: {file_write}')

    compare_files(file_fullname, decrypted_file)
    copy_files_for_matica(file_fullname)
    automate_maticard(file_write, file_fullname)
    delete_file(decrypted_file)

main()
