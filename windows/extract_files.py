import os
import subprocess
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

def remove_non_ascii(input_string):
    # Remove all characters outside the ASCII range
    clean_string = ''.join(char for char in input_string if 0 < ord(char) < 128)
    return clean_string


def list_compressed_content(folder_name,archive_file):
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
                    path_value = remove_non_ascii(current_line.split('=')[1].strip())
                    paths_list.append(path_value)

            return set(paths_list), lines
        else:
            error_message = f"Error (return code {result.returncode}): {result.stderr}\n"
            error_log = open(folder_name + "/listing_error.txt", "a")
            error_log.write(error_message +"\n")
            return set(["Comando de listagem retornou valor diferente de 0"])
    except Exception as e:
        error_log = open(folder_name + "/listing_error.txt", "a")
        error_log.write(f"An error occurred: {str(e)} on {str(archive_file)}\n")
        return set(["Erro na listagem dos arquivos comprimidos"])
    
def list_files(folder):
    files_list = []

    for folder_path, _, files in os.walk(folder):
        for file_name in files:
            file_path = os.path.join(folder_path, file_name)
            
            # Obtém o caminho relativo ao diretório inicial
            relative_path = os.path.relpath(file_path, folder)
            relative_path = remove_non_ascii(relative_path)
            files_list.append(relative_path)

    return set(files_list)

def autorename_check(missing_files, surplus_files):
    renamed_list=[]
    for path in surplus_files:
        if '_1.' in path:
            renamed = path.replace('_1.', '.')
            renamed_list.append(renamed)
    new_missing=missing_files.difference(set(renamed_list))
    return new_missing

def compare_content(folder_name,archive_file, output_dir):
    a=list_compressed_content(folder_name,archive_file)
    if(len(a)>1):
        before_extraction, lines=a[0], a[1]
    else:
        before_extraction=a
    after_extraction=list_files(output_dir)
    differences_both = before_extraction.symmetric_difference(after_extraction)
    missing_files=before_extraction.difference(after_extraction)
    surplus_files=after_extraction.difference(before_extraction)
    file_name=os.path.splitext(os.path.basename(archive_file))[0]
    new_missing=autorename_check(missing_files,surplus_files)

    if(len(new_missing)>0):
        os.makedirs(folder_name + "/errors", exist_ok=True)
        output_file_path = folder_name + "/errors/"+file_name +".txt"
        with open(output_file_path, 'w') as file:
            file.write(f"Numero de items de antes que nao tem agora {len(missing_files)}\n\n\n\n")
            file.write(f"Items que tinham antes mas nao tem depois: {missing_files}\n\n\n\n\n")
            file.write(f"Items que surgiram depois mas nao tem antes: {surplus_files}\n\n\n\n\n")
            file.write(f"\n\n\n\n\n\nCAMINHO ANTES: {archive_file}\n")
            file.write(f"Files before {len(before_extraction)} \n")
            file.write(str(before_extraction)+"\n")
            file.write(f"\n\n\n\n\n\nCAMINHO DEPOIS: {output_dir}\n")
            file.write(f"Files after {len(after_extraction)} \n")
            file.write(str(after_extraction)+"\n")
            for difference in differences_both:
                file.write(f"Item {difference} is in one set but not in the other\n")
            if(len(a)>1):
                for line in lines:
                    file.write(line +"\n")
        return False

    return True



# Função que extrai arquivos e registra caso ocorra um erro ou um sucesso
def extract_archive(folder_name,archive_file, output_dir, auto_rename, remove_zips, yes_toall):
    _, extension = os.path.splitext(archive_file)
    supported_extensions = ['.zip', '.rar', '.7z']
    
    if extension.lower() in supported_extensions:
        os.makedirs(output_dir, exist_ok=True)
        command = [SETE_ZIP_PATH, "x", archive_file, f"-o{output_dir}"]
        if yes_toall:
            command.append("-y")
        if auto_rename:
            command.append("-aou")
    else:
        return

    try:
        extract_log = open(folder_name + "/extract_log.txt", "a")
        subprocess.run(command, check=True)
        if(compare_content(folder_name,archive_file, output_dir)==True):
            extract_log.write(f"Extraction successful: {archive_file}\n")
            extract_log.close()
            if not remove_zips:
                os.remove(archive_file)
        else:
            error_log = open(folder_name + "/extract_error.txt", "a")
            error_log.write(f"Extraction failed: {archive_file}\n")
            error_log.write("Error on comparing zip content, check file error\n")


    except subprocess.CalledProcessError as e:
        error_log = open(folder_name + "/extract_error.txt", "a")
        error_log.write(f"Extraction failed: {archive_file}\n")
        error_log.write(str(e) + "\n")
        tam = len(archive_file)
        if tam > 256:
            error_log.write(f"O caminho do arquivo tem tamanho: {str(tam)}, você deve reduzir {str(tam - 256)} caracteres \n")
        error_log.close()


# Função auxiliar de extração recursiva
def extract_recursive(folder_name,file_path, checkbox_rename_var, checkbox_zip_var, checkbox_yes_var):
    output_folder = os.path.splitext(os.path.basename(file_path))[0]
    output_folder = os.path.dirname(file_path)+"/"+output_folder.replace(" ", "_")
    extract_archive(folder_name,file_path, output_folder, checkbox_rename_var, checkbox_zip_var, checkbox_yes_var)

    for root, dirs, files in os.walk(output_folder):
        for file in files:
            all_files=open(folder_name+"/file_listing/all_extracted_files.txt","a")
            files256=open(folder_name+"/file_listing/nomes_maiores_256.txt","a")
            file_path = os.path.join(root, file)
            tam=len(file_path)
            all_files.write(str(tam)+" caracteres: "+file_path+"\n")
            if(tam>256): files256.write(str(tam)+" caracteres: "+file_path+"\n")
            extract_recursive(folder_name,file_path, checkbox_rename_var, checkbox_zip_var, checkbox_yes_var)


# Funções para o frontend
            
class HoverInfo:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.widget.bind("<Enter>", self.on_enter)
        self.widget.bind("<Leave>", self.on_leave)

    def on_enter(self, event=None):
        self.tooltip = tk.Toplevel(self.widget)
        x, y, _, _ = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 25
        self.tooltip.wm_overrideredirect(True)
        self.tooltip.wm_geometry(f"+{x}+{y}")
        tk.Label(self.tooltip, text=self.text, background="#ffffe0", relief="solid", borderwidth=1).grid()

    def on_leave(self, event=None):
        if self.tooltip:
            self.tooltip.destroy()


# Botão extrair
def on_extrair_click():
    error_count = 0
    i = 0
    progress_var.set(0)
    progress_bar["style"] = "sucess.Horizontal.TProgressbar"
    progress_bar.update()
    for file_path in listbox_extract.get(0, tk.END):
        folder_name = "extraction_logs/results_" + os.path.splitext(os.path.basename(file_path))[0]
        os.makedirs(folder_name, exist_ok=True)
        os.makedirs(folder_name+"/file_listing", exist_ok=True)
        extract_recursive(folder_name,file_path, checkbox_rename_var.get(), checkbox_zip_var.get(), checkbox_yes_var.get())
        i += int(100 / listbox_extract.size())
        progress_var.set(i)
        progress_bar.update()

        # Verifica se ocorreu algum erro na extração
        error_log_path = folder_name + "/extract_error.txt"
        if os.path.exists(error_log_path):
            with open(error_log_path, "r") as error_log:
                errors = error_log.read()
                error_count += errors.count("Extraction failed")
                progress_bar["style"] = "error.Horizontal.TProgressbar"
                progress_bar.update()

    # Altera a cor da barra de progresso com base na contagem de erros
    if error_count > 0:
        progress_bar["style"] = "error.Horizontal.TProgressbar"
        progress_bar.update()
        error_label.config(text=f"Número de erros: {error_count}")
        messagebox.showinfo("Extração Concluída com erros", "Cheque os logs do processo disponíveis na pasta extraction_logs")
    else:
        progress_bar["style"] = "sucess.Horizontal.TProgressbar"
        progress_bar.update()
        error_label.config(text="Número de erros: 0")
        messagebox.showinfo("Extração Concluída", "ERRO: Os logs do processo estão disponíveis na pasta extraction_logs")



# Caixas de Seleção
def on_checkbox_change():
    auto_rename = checkbox_rename_var.get()
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
    checkbox_rename_var.set(True)
    checkbox_yes_var.set(True)
    checkbox_zip_var.set(False)
    progress_var.set(0)
    label_count.config(text=f"Número de arquivos: 0")


def remove_item():
    selection = listbox_extract.curselection()
    if selection:
        index = selection[0]
        listbox_extract.delete(index)

# Criação dos logs
os.makedirs("extraction_logs", exist_ok=True)

# Criação da Interface
root = tk.Tk()
s = ttk.Style()
s.theme_use('clam')
s.configure("error.Horizontal.TProgressbar", troughcolor="white", background="green",thickness=15)
s.configure("sucess.Horizontal.TProgressbar", troughcolor="white", background="green",thickness=15)

root.columnconfigure(0, weight=1)
root.columnconfigure(1, weight=1)

listbox_count = tk.Listbox(root, selectmode=tk.SINGLE, width=50, height=5)
listbox_count.grid(pady=10, row=0, columnspan=2)

# Listador de Arquivos Comprimidos
label_count = tk.Label(root, text="Número de arquivos: 0")
label_count.grid(pady=5, columnspan=2, row=1)



frame_list = tk.Frame(root)
frame_list.grid(row=2, column=0, sticky="nsew") 

button_listar_arquivos = tk.Button(frame_list, text="Listar Arquivos Comprimidos", command=on_listar_arquivos_click)
button_listar_arquivos.grid(padx=40,pady=10, sticky='ew')

# Tooltip para o botão Listar Arquivos Comprimidos
HoverInfo(button_listar_arquivos, "Selecione pasta(s) para listar todos os arquivos comprimidos presentes nela(s)")

# Botão Copiar

frame_copy = tk.Frame(root)
frame_copy.grid(row=2, column=1, sticky="nsew") 


button_copy = tk.Button(frame_copy, text="Copiar todos", command=on_copy_click)
button_copy.grid(pady=5, sticky='ew')
HoverInfo(button_copy, "Copia a lista de arquivos comprimidos acima para a caixa abaixo")

# Descompactador de Arquivos
listbox_extract = tk.Listbox(root, selectmode=tk.MULTIPLE, width=50, height=5)
listbox_extract.grid(pady=10, row=3, columnspan=2)



frame_select = tk.Frame(root)
frame_select.grid(row=4, column=0, sticky="nsew") 


button_selecionar = tk.Button(frame_select, text="Selecionar Arquivos", command=on_select_click)
button_selecionar.grid(padx=40,pady=10, sticky='ew')
HoverInfo(button_selecionar, "Selecione arquivo(s) comprimidos")


frame_remove = tk.Frame(root)
frame_remove.grid(row=4, column=1, sticky="nsew") 

button_remover = tk.Button(frame_remove, text="Remover Arquivo", command=remove_item)
button_remover.grid(pady=10, sticky='ew')
HoverInfo(button_remover, "Remover arquivo selecionado")


# Caixas de Seleção
checkbox_rename_var = tk.BooleanVar(value=True)
checkbox_rename = tk.Checkbutton(root, text="Auto Renomear arquivos com nomes repetidos", variable=checkbox_rename_var,
                                  command=on_checkbox_change)
checkbox_rename.grid(pady=5,row=5, columnspan=2)
HoverInfo(checkbox_rename,"Automaticamente renomeia arquivos com nomes repetidos durante a extração")

checkbox_zip_var = tk.BooleanVar()
checkbox_zip = tk.Checkbutton(root, text="Não apagar arquivos comprimidos", variable=checkbox_zip_var,
                              command=on_checkbox_change)
checkbox_zip.grid(pady=5, row=6,columnspan=2)
HoverInfo(checkbox_zip,"Extrai os arquivos sem remover os arquivos comprimidos")

checkbox_yes_var = tk.BooleanVar(value=True)
checkbox_yes = tk.Checkbutton(root, text="Sim para todas as perguntas", variable=checkbox_yes_var,
                              command=on_checkbox_change)
checkbox_yes.grid(pady=5, row=7,columnspan=2)
HoverInfo(checkbox_yes,"Responde sim para todas as perguntas do 7zip durante a extração")

# Barra de Progresso da Descompressão
#style = ttk.Style()
#tyle.configure("TProgressbar", thickness=15)

progress_var = tk.IntVar()

progress_bar = ttk.Progressbar(root, variable=progress_var, mode='determinate', length=320)
progress_bar.grid(pady=5, row=8, columnspan=2)

#style.configure("success.Horizontal.TProgressbar", troughcolor="white", background="green", thickness=15)
#style.configure("error.Horizontal.TProgressbar", troughcolor="white", background="yellow", thickness=15) 

# Label para exibir o número de erros
error_label = tk.Label(root, text="Número de erros: 0",bd=0, highlightthickness=0)
error_label.grid(pady=5, row=9,columnspan=2)


frame_extrair = tk.Frame(root)
frame_extrair.grid(row=10, column=1, sticky="nsew") 

button_extrair = tk.Button(frame_extrair, text="Extrair", command=on_extrair_click)
button_extrair.grid(pady=10,row=10, column=1, sticky="ew")
HoverInfo(button_extrair, "Extraia os arquivos listados na caixa acima")


frame_reset = tk.Frame(root)
frame_reset.grid(row=10, column=0, sticky="nsew") 

# Botão de Reiniciar
button_reiniciar = tk.Button(frame_reset, text="Reiniciar Programa", command=reiniciar_programa)
button_reiniciar.grid(padx=40,pady=10, row=10, column=0)
HoverInfo(button_reiniciar, "Reinicia o programa e limpa as caixas")

root.geometry("400x600")
root.title("Descompactação de Arquivos Geo Ansata")
root.mainloop()
