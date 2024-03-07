import os
import subprocess
import datetime
import tkinter as tk
from tkinter import filedialog, ttk

#Criação dos logs
os.makedirs("extraction_logs", exist_ok=True)
current_datetime = datetime.datetime.now()
timestamp_format = "%d%m%y%H%M%S"
folder_name = "extraction_logs/results_"+current_datetime.strftime(timestamp_format)
os.makedirs(folder_name, exist_ok=True)



#Função que extrai arquivos e registra caso ocorra um erro ou um sucesso
def extract_archive(archive_file, output_dir, auto_convert_unicode, remove_zips):
    # Determine the archive format based on the file extension
    _, extension = os.path.splitext(archive_file)    
    supported_extensions = ['.zip', '.rar', '.7z']

    if extension.lower() in supported_extensions:
        os.makedirs(output_dir, exist_ok=True)
        command = ["7z", "x", archive_file, f"-o{output_dir}"]
        # Add switches if enabled
        if auto_convert_unicode:
            command.append("-aou")

        print(command)
    else:
        return

    try:
        # Run the appropriate 7zip command
        extract_log = open(folder_name + "/extract_log.txt", "a")
        subprocess.run(command, check=True)
        extract_log.write(f"Extraction successful: {archive_file}\n")
        extract_log.close()
        # Remove the original compressed file
        if not remove_zips:
            os.remove(archive_file)
    except subprocess.CalledProcessError as e:
        error_log = open(folder_name + "/extract_error.txt", "a")
        error_log.write(f"Extraction failed: {archive_file}\n")
        error_log.write(str(e)+"\n")
        tam=len(archive_file)
        if(tam>256):
            error_log.write(f"O caminho do arquivo tem tamanho: {str(tam)}, você deve reduzir {str(tam - 256)} caracteres \n")
        error_log.close()

        
#Função auxiliar de extração recursiva
def extract_recursive(zip_file,checkbox_unicode_var,checkbox_zip_var):
    # Extract the file to the corresponding folder
    output_folder = os.path.splitext(zip_file)[0]
    extract_archive(zip_file, output_folder,checkbox_unicode_var, checkbox_zip_var)

    # Check for compressed files within the extracted folder
    for root, dirs, files in os.walk(output_folder):
        for file in files:
            file_path = os.path.join(root, file)

            # Recursively process compressed files within the extracted folder
            extract_recursive(file_path,checkbox_unicode_var,checkbox_zip_var)        



#Funções para o frontend
def on_descompactar_click():
    error_count = 0
    i=0
    for file_path in listbox_extract.get(0, tk.END):
        extract_recursive(file_path,checkbox_unicode_var.get(),checkbox_zip_var.get())
        i+=int(100/listbox_extract.size())
        progress_var.set(i)  # Update the progress variable
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




def on_checkbox_change():
    auto_convert_unicode = checkbox_unicode_var.get()
    remove_zips=checkbox_zip_var.get()




def on_select_click():
    file_paths = filedialog.askopenfilenames(
        title="Selecione Arquivos",
        filetypes=[
            ("Arquivos de Texto", "*.7z"),
            ("Arquivos de Texto", "*.zip"),
            ("Arquivos de Texto", "*.rar"),
            ("Todos os Arquivos", "*.*")
        ],
        initialdir="/media/gabriel/SSD/GeoAnsata"
    )
    for file_path in file_paths:
        listbox_extract.insert(tk.END, file_path)




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
        

def on_listar_arquivos_click():
    listbox_count.delete(0, tk.END)  # Limpar a Listbox antes de listar novamente
    folder_selected = filedialog.askdirectory(title="Selecione uma pasta")
    if folder_selected:
        list_and_count_compressed_files(folder_selected, listbox_count, label_count)



root = tk.Tk()


label_count = tk.Label(root, text="Número de arquivos: ?")
label_count.pack(pady=5)

listbox_count = tk.Listbox(root, selectmode=tk.SINGLE, width=50, height=10)
listbox_count.pack(pady=10)

button_listar_arquivos = tk.Button(root, text="Listar Arquivos Comprimidos", command=on_listar_arquivos_click)
button_listar_arquivos.pack(pady=10)




listbox_extract = tk.Listbox(root, selectmode=tk.MULTIPLE, width=50, height=10)
listbox_extract.pack(pady=10)

button_selecionar = tk.Button(root, text="Selecionar Arquivos", command=on_select_click)
button_selecionar.pack(pady=10)

button_descompactar = tk.Button(root, text="Descompactar", command=on_descompactar_click)
button_descompactar.pack(pady=5)

checkbox_unicode_var = tk.BooleanVar()
checkbox_unicode = tk.Checkbutton(root, text="Auto Converter para Unicode", variable=checkbox_unicode_var, command=on_checkbox_change)
checkbox_unicode.pack(pady=5)


checkbox_zip_var = tk.BooleanVar()
checkbox_zip = tk.Checkbutton(root, text="Não apagar arquivos comprimidos", variable=checkbox_zip_var, command=on_checkbox_change)
checkbox_zip.pack(pady=5)


# Create a progress variable
style = ttk.Style()
style.configure("TProgressbar", thickness=20)

progress_var = tk.IntVar()

progress_bar = ttk.Progressbar(root, variable=progress_var, mode='determinate')
progress_bar.pack(pady=15)

style.configure("success.Horizontal.TProgressbar", troughcolor="white", background="green", thickness=15)
style.configure("error.Horizontal.TProgressbar", troughcolor="white", background="yellow", thickness=15)

# Label para exibir o número de erros
error_label = tk.Label(root, text="")
error_label.pack(pady=5)

root.geometry("400x600")
root.title("Descompactação de Arquivos Geo Ansata")
root.mainloop()