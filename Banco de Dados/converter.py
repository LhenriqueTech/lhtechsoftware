import os
import tkinter as tk
from tkinter import filedialog, messagebox
from concurrent.futures import ThreadPoolExecutor

def convert_file(input_path, output_path):
    """Função para converter um arquivo de UTF-8 para ANSI."""
    try:
        with open(input_path, 'r', encoding='utf-8') as infile:
            content = infile.read()

        with open(output_path, 'w', encoding='cp1252') as outfile:
            outfile.write(content)

        print(f"Arquivo {os.path.basename(input_path)} convertido com sucesso para ANSI.")
    except Exception as e:
        print(f"Erro ao converter o arquivo {os.path.basename(input_path)}: {e}")

def convert_utf8_to_ansi(input_folder, output_folder):
    """Função para converter todos os arquivos de uma pasta de UTF-8 para ANSI."""
    if not os.path.exists(input_folder) or not os.path.exists(output_folder):
        print("Erro: uma das pastas fornecidas não existe.")
        return

    txt_files = [f for f in os.listdir(input_folder) if f.endswith(".txt")]
    
    if not txt_files:
        print("Nenhum arquivo .txt encontrado.")
        return

    # Usar ThreadPoolExecutor para paralelizar o processo
    with ThreadPoolExecutor() as executor:
        for filename in txt_files:
            input_path = os.path.join(input_folder, filename)
            output_path = os.path.join(output_folder, "ANSI_" + filename)
            executor.submit(convert_file, input_path, output_path)

    print("Conversão finalizada para todos os arquivos.")

def select_input_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        input_folder_var.set(folder_selected)

def select_output_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        output_folder_var.set(folder_selected)

def start_conversion():
    input_folder = input_folder_var.get()
    output_folder = output_folder_var.get()

    if not input_folder or not output_folder:
        messagebox.showerror("Erro", "Por favor, selecione as pastas de entrada e saída.")
        return

    convert_utf8_to_ansi(input_folder, output_folder)
    messagebox.showinfo("Sucesso", "Conversão concluída com sucesso!")

# Configuração da interface gráfica
root = tk.Tk()
root.title("Conversor UTF-8 para ANSI")

input_folder_var = tk.StringVar()
output_folder_var = tk.StringVar()

tk.Label(root, text="Pasta de Entrada:").grid(row=0, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=input_folder_var, width=40).grid(row=0, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=select_input_folder).grid(row=0, column=2, padx=10, pady=10)

tk.Label(root, text="Pasta de Saída:").grid(row=1, column=0, padx=10, pady=10)
tk.Entry(root, textvariable=output_folder_var, width=40).grid(row=1, column=1, padx=10, pady=10)
tk.Button(root, text="Selecionar", command=select_output_folder).grid(row=1, column=2, padx=10, pady=10)

tk.Button(root, text="Iniciar Conversão", command=start_conversion).grid(row=2, column=0, columnspan=3, padx=10, pady=20)

root.mainloop()
