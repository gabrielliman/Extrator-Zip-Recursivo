import os
import subprocess
import datetime
import tkinter as tk
from tkinter import filedialog, ttk, messagebox
from unidecode import unidecode
SETE_ZIP_PATH = 'C:\\Program Files\\7-Zip\\7z.exe'

def remove_broken_characters(input_string):
    try:
        # Try to encode the string using UTF-8 and then decode it
        clean_string = input_string.encode('utf-8').decode('utf-8')
    except UnicodeDecodeError:
        # If encoding/decoding fails, replace broken characters with an empty string
        clean_string = ''.join(char for char in input_string if ord(char) < 128)

    return clean_string



def std_text(texto):
    # Dicionário de substituições para acentos
    substituicoes_acentos = {
        'á': 'a', 'à': 'a', 'â': 'a', 'ã': 'a',
        'é': 'e', 'è': 'e', 'ê': 'e',
        'í': 'i', 'ì': 'i', 'î': 'i',
        'ó': 'o', 'ò': 'o', 'ô': 'o', 'õ': 'o',
        'ú': 'u', 'ù': 'u', 'û': 'u',
        'ç': 'c'
    }

    # Substitui acentos e cedilha
    for acento, sem_acento in substituicoes_acentos.items():
        texto = texto.replace(acento, sem_acento)

    return texto



def list_compressed_content(archive_file):
    try:
        command = [SETE_ZIP_PATH, "l", archive_file, "-ba","-slt"]
        result = subprocess.run(command, stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True)

        if result.returncode == 0:
            paths_list = []
            lines = result.stdout.split('\n')
            
            for i in range(len(lines) - 1):
                current_line = lines[i]
                next_line = lines[i + 1]

                if current_line.startswith('Path') and next_line.startswith('Folder = -'):
                    path_value = std_text(current_line.split('=')[1].strip())
                    path_value = std_text(path_value)
                    paths_list.append(path_value)

            return set(paths_list)
        else:
            error_message = f"Error (return code {result.returncode}): {result.stderr}\n"
            error_log = open(folder_name + "/listing_error.txt", "a")
            error_log.write(error_message)
            return error_message
    except Exception as e:
        error_log = open(folder_name + "/listing_error.txt", "a")
        error_log.write(f"An error occurred: {str(e)}")
        return f"An error occurred: {str(e)}"
    
def list_files(folder):
    files_list = []

    for folder_path, _, files in os.walk(folder):
        for file_name in files:
            file_path = os.path.join(folder_path, file_name)
            
            # Obtém o caminho relativo ao diretório inicial
            relative_path = os.path.relpath(file_path, folder)
            relative_path = std_text(relative_path)
            
            files_list.append(relative_path)

    return set(files_list)

def compare_content(archive_file, output_dir):
    before_extraction=list_compressed_content(archive_file)
    after_extraction=list_files(output_dir)
    differences_both = before_extraction.symmetric_difference(after_extraction)
    missing_files=before_extraction.difference(after_extraction)
    file_name=os.path.splitext(os.path.basename(archive_file))[0]
    output_file_path = folder_name + file_name +".txt"


    if(len(missing_files)>0):
        output_file_path = folder_name + "/"+file_name +".txt"
        with open(output_file_path, 'w') as file:
            file.write(f"Numero de items de antes que nao tem agora {len(missing_files)}\n\n\n\n")
            file.write(f"Items que tinham antes mas nao tem depois: {before_extraction.difference(after_extraction)}\n\n\n\n\n")
            file.write(f"Items que surgiram depois mas nao tem antes: {after_extraction.difference(before_extraction)}\n\n\n\n\n")
            file.write(f"\n\n\n\n\n\nCAMINHO ANTES: {archive_file}\n")
            file.write(f"Files before {len(before_extraction)} \n")
            file.write(str(before_extraction)+"\n")
            file.write(f"\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\n\nCAMINHO DEPOIS: {output_dir}\n")
            file.write(f"Files after {len(after_extraction)} \n")
            file.write(str(after_extraction)+"\n")
            for difference in differences_both:
                file.write(f"Item {difference} is in one set but not in the other\n")

    return before_extraction==after_extraction



# Função que extrai arquivos e registra caso ocorra um erro ou um sucesso
def extract_archive(archive_file, output_dir, auto_convert_unicode, remove_zips, yes_toall):
    _, extension = os.path.splitext(archive_file)
    supported_extensions = ['.zip', '.rar', '.7z']
    
    if extension.lower() in supported_extensions:
        os.makedirs(output_dir, exist_ok=True)
        command = [SETE_ZIP_PATH, "x", archive_file, f"-o{output_dir}"]
        if yes_toall:
            command.append("-y")
        if auto_convert_unicode:
            command.append("-aou")
    else:
        return

    try:
        extract_log = open(folder_name + "/extract_log.txt", "a")
        subprocess.run(command, check=True)
        if(compare_content(archive_file, output_dir)==True):
            extract_log.write(f"Extraction successful: {archive_file}\n")
            extract_log.close()
            if not remove_zips:
                os.remove(archive_file)
        else:
            error_log = open(folder_name + "/extract_error.txt", "a")
            error_log.write(f"Extraction failed: {archive_file}\n")
            error_log.write("Files missing\n")


    except subprocess.CalledProcessError as e:
        error_log = open(folder_name + "/extract_error.txt", "a")
        error_log.write(f"Extraction failed: {archive_file}\n")
        error_log.write(str(e) + "\n")
        tam = len(archive_file)
        if tam > 256:
            error_log.write(f"O caminho do arquivo tem tamanho: {str(tam)}, você deve reduzir {str(tam - 256)} caracteres \n")
        error_log.close()


# Função auxiliar de extração recursiva
def extract_recursive(zip_file, checkbox_unicode_var, checkbox_zip_var, checkbox_yes_var):
    output_folder = os.path.splitext(zip_file)[0]
    output_folder = output_folder.replace(" ", "_")
    extract_archive(zip_file, output_folder, checkbox_unicode_var, checkbox_zip_var, checkbox_yes_var)

    for root, dirs, files in os.walk(output_folder):
        for file in files:
            file_path = os.path.join(root, file)
            extract_recursive(file_path, checkbox_unicode_var, checkbox_zip_var, checkbox_yes_var)


# Funções para o frontend

# Botão Descompactar
def on_descompactar_click():
    error_count = 0
    i = 0
    for file_path in listbox_extract.get(0, tk.END):
        extract_recursive(file_path, checkbox_unicode_var.get(), checkbox_zip_var.get(), checkbox_yes_var.get())
        i += int(100 / listbox_extract.size())
        progress_var.set(i)
        progress_bar.update()

        # Verifica se ocorreu algum erro na extração
        error_log_path = folder_name + "/extract_error.txt"
        if os.path.exists(error_log_path):
            with open(error_log_path, "r") as error_log:
                errors = error_log.read()
                error_count += errors.count("Extraction failed")

    # Altera a cor da barra de progresso com base na contagem de erros
    if error_count > 0:
        progress_bar["style"] = "error.Horizontal.TProgressbar"
        error_label.config(text=f"Número de erros: {error_count}")
    else:
        progress_bar["style"] = "success.Horizontal.TProgressbar"
        error_label.config(text="")
    messagebox.showinfo("Extração Concluída", "Os logs do processo estão disponíveis na pasta extraction_logs")



# Caixas de Seleção
def on_checkbox_change():
    auto_convert_unicode = checkbox_unicode_var.get()
    remove_zips = checkbox_zip_var.get()
    yes_toall = checkbox_yes_var.get()


# Botão Selecionar Arquivos
def on_select_click():
    file_paths = filedialog.askopenfilenames(
        title="Selecione Arquivos",
        filetypes=[
            ("Arquivos de Texto", "*.7z"),
            ("Arquivos de Texto", "*.zip"),
            ("Arquivos de Texto", "*.rar"),
            ("Todos os Arquivos", "*.*")
        ],
        initialdir="./"
    )
    for file_path in file_paths:
        listbox_extract.insert(tk.END, file_path)


# Função Auxiliar para Listar Arquivos Comprimidos
def list_and_count_compressed_files(folder_selected, listbox_count, label_count):
    supported_extensions = ['.zip', '.rar', '.7z']
    if folder_selected:
        for root, dirs, files in os.walk(folder_selected):
            for file in files:
                file_path = os.path.join(root, file)
                _, extension = os.path.splitext(file_path)

                if extension.lower() in supported_extensions:
                    listbox_count.insert(tk.END, file_path)

        file_count = listbox_count.size()
        label_count.config(text=f"Número de arquivos: {file_count}")


# Botão Listar Arquivos Comprimidos
def on_listar_arquivos_click():
    listbox_count.delete(0, tk.END)  # Limpar a Listbox antes de listar novamente
    folder_selected = filedialog.askdirectory(title="Selecione uma pasta")
    if folder_selected:
        list_and_count_compressed_files(folder_selected, listbox_count, label_count)


# Função que copia de uma lista para outra
def on_copy_click():
    selected_items = listbox_count.curselection()
    for item in listbox_count.get(0, tk.END):
        listbox_extract.insert(tk.END, item)


# Botão Reiniciar Programa
def reiniciar_programa():
    # Limpar a interface e redefinir variáveis
    listbox_count.delete(0, tk.END)
    listbox_extract.delete(0, tk.END)
    error_label.config(text="")
    checkbox_unicode_var.set(False)
    checkbox_yes_var.set(True)
    checkbox_zip_var.set(False)
    progress_var.set(0)

    # Criar novos logs
    global folder_name
    current_datetime = datetime.datetime.now()
    timestamp_format = "%H-%M-%S-%d-%m-%y"
    folder_name = "extraction_logs/results_" + current_datetime.strftime(timestamp_format)
    os.makedirs(folder_name, exist_ok=True)


# Criação dos logs
os.makedirs("extraction_logs", exist_ok=True)
current_datetime = datetime.datetime.now()
timestamp_format = "%H-%M-%S-%d-%m-%y"
folder_name = "extraction_logs/results_" + current_datetime.strftime(timestamp_format)
os.makedirs(folder_name, exist_ok=True)

# Criação da Interface
root = tk.Tk()

# Listador de Arquivos Comprimidos
label_count = tk.Label(root, text="Número de arquivos: ?")
label_count.pack(pady=5)

listbox_count = tk.Listbox(root, selectmode=tk.SINGLE, width=50, height=10)
listbox_count.pack(pady=10)

button_listar_arquivos = tk.Button(root, text="Listar Arquivos Comprimidos", command=on_listar_arquivos_click)
button_listar_arquivos.pack(pady=5)

# Copiador de Lista
button_copy = tk.Button(root, text="Copiar para Descompactar", command=on_copy_click)
button_copy.pack(pady=5)

# Descompactador de Arquivos
listbox_extract = tk.Listbox(root, selectmode=tk.MULTIPLE, width=50, height=10)
listbox_extract.pack(pady=10)

button_selecionar = tk.Button(root, text="Selecionar Arquivos", command=on_select_click)
button_selecionar.pack(pady=10)

button_descompactar = tk.Button(root, text="Descompactar", command=on_descompactar_click)
button_descompactar.pack(pady=5)

# Caixas de Seleção
checkbox_unicode_var = tk.BooleanVar()
checkbox_unicode = tk.Checkbutton(root, text="Auto Converter para Unicode", variable=checkbox_unicode_var,
                                  command=on_checkbox_change)
checkbox_unicode.pack(pady=5)

checkbox_zip_var = tk.BooleanVar()
checkbox_zip = tk.Checkbutton(root, text="Não apagar arquivos comprimidos", variable=checkbox_zip_var,
                              command=on_checkbox_change)
checkbox_zip.pack(pady=5)

checkbox_yes_var = tk.BooleanVar(value=True)
checkbox_yes = tk.Checkbutton(root, text="Sim para todas as perguntas", variable=checkbox_yes_var,
                              command=on_checkbox_change)
checkbox_yes.pack(pady=5)

# Barra de Progresso da Descompressão
style = ttk.Style()
style.configure("TProgressbar", thickness=20)

progress_var = tk.IntVar()

progress_bar = ttk.Progressbar(root, variable=progress_var, mode='determinate', length=400)
progress_bar.pack(pady=15)

style.configure("success.Horizontal.TProgressbar", troughcolor="white", background="green", thickness=15)
style.configure("error.Horizontal.TProgressbar", troughcolor="white", background="yellow", thickness=15)

# Label para exibir o número de erros
error_label = tk.Label(root, text="")
error_label.pack(pady=5)

# Botão de Reiniciar
button_reiniciar = tk.Button(root, text="Reiniciar Programa", command=reiniciar_programa)
button_reiniciar.pack(pady=5)

root.geometry("600x800")
root.title("Descompactação de Arquivos Geo Ansata")
root.mainloop()
