import tkinter as tk
from tkinter import filedialog, messagebox
from docx2pdf import convert
import os

def select_docx_file():
    file_path = filedialog.askopenfilename(
        title="Selecione o arquivo .docx",
        filetypes=[("Documentos Word", "*.docx")]
    )
    docx_entry.delete(0, tk.END)
    docx_entry.insert(0, file_path)

def select_output_folder():
    folder_path = filedialog.askdirectory(
        title="Selecione a pasta de saída"
    )
    output_entry.delete(0, tk.END)
    output_entry.insert(0, folder_path)

def convert_to_pdf():
    docx_file = docx_entry.get()
    output_folder = output_entry.get()

    if not docx_file or not output_folder:
        messagebox.showerror("Erro", "Por favor, selecione o arquivo .docx e a pasta de saída.")
        return

    try:
        # Realiza a conversão
        output_pdf = os.path.join(output_folder, os.path.splitext(os.path.basename(docx_file))[0] + ".pdf")
        convert(docx_file, output_pdf)
        messagebox.showinfo("Sucesso", f"Conversão concluída! PDF salvo em:\n{output_pdf}")
    except Exception as e:
        messagebox.showerror("Erro", f"Falha na conversão: {e}")

# Configuração da interface gráfica
root = tk.Tk()
root.title("Conversor de .docx para PDF")

# Widgets de entrada do arquivo .docx
docx_label = tk.Label(root, text="Arquivo .docx:")
docx_label.grid(row=0, column=0, padx=10, pady=10, sticky="e")

docx_entry = tk.Entry(root, width=40)
docx_entry.grid(row=0, column=1, padx=10, pady=10)

docx_button = tk.Button(root, text="Selecionar", command=select_docx_file)
docx_button.grid(row=0, column=2, padx=10, pady=10)

# Widgets de seleção da pasta de saída
output_label = tk.Label(root, text="Pasta de saída:")
output_label.grid(row=1, column=0, padx=10, pady=10, sticky="e")

output_entry = tk.Entry(root, width=40)
output_entry.grid(row=1, column=1, padx=10, pady=10)

output_button = tk.Button(root, text="Selecionar", command=select_output_folder)
output_button.grid(row=1, column=2, padx=10, pady=10)

# Botão de conversão
convert_button = tk.Button(root, text="Converter para PDF", command=convert_to_pdf)
convert_button.grid(row=2, column=0, columnspan=3, pady=20)

# Execução da interface
root.mainloop()
