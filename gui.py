import tkinter as tk
from tkinter import filedialog, messagebox
from main import extrair_texto_pdfs

class ExtratorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Extrator de Texto de PDFs")
        self.root.geometry("300x300")
        
        self.root.configure(bg="White")

        self.lista_arquivos = []

        self.label_info = tk.Label(root, text="Nenhum arquivo selecionado.")
        self.label_info.pack(pady=10)

        self.btn_selecionar = tk.Button(root, text="Selecionar PDFs", command=self.selecionar_pdfs)
        self.btn_selecionar.pack(pady=10)

        self.btn_extrair = tk.Button(root, text="Extrair Texto", command=self.extrair_texto)
        self.btn_extrair.pack(pady=10)

    def selecionar_pdfs(self):
        arquivos = filedialog.askopenfilenames(title="Selecione os arquivos PDF", filetypes=[("Arquivos PDF", "*.pdf")])
        if arquivos:
            self.lista_arquivos = list(arquivos)
            self.label_info.config(text=f"{len(self.lista_arquivos)} arquivo(s) selecionado(s).")
        else:
            self.label_info.config(text="Nenhum arquivo selecionado.")

    def extrair_texto(self):
        if not self.lista_arquivos:
            messagebox.showwarning("Aviso", "Nenhum PDF selecionado.")
            return

        caminho_saida = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Arquivo de texto", "*.txt")])
        if not caminho_saida:
            return

        sucesso, mensagem = extrair_texto_pdfs(self.lista_arquivos, caminho_saida)
        if sucesso:
            messagebox.showinfo("Sucesso", mensagem)
        else:
            messagebox.showerror("Erro", mensagem)

if __name__ == "__main__":
    root = tk.Tk()
    app = ExtratorGUI(root)
    root.mainloop()
