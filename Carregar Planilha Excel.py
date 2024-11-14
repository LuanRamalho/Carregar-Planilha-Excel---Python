import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

class ExcelReaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Leitor de Excel")
        self.root.geometry("600x400")
        self.root.configure(bg="#4CAF50")

        # Estilo do Treeview
        style = ttk.Style()
        style.configure("Treeview", foreground="black", font=("Arial", 12))
        style.configure("Treeview.Heading", background="#81C784", font=("Arial", 14, "bold"))

        # Label de instruções
        self.label = tk.Label(self.root, text="Escolha um arquivo Excel para ler:", bg="#4CAF50", font=("Arial", 14))
        self.label.pack(pady=20)

        # Botão para abrir arquivo
        self.open_button = tk.Button(self.root, text="Abrir Arquivo", command=self.load_file, bg="#FFEB3B", font=("Arial", 12))
        self.open_button.pack(pady=10)

        # Frame para o Treeview
        self.frame = ttk.Frame(self.root)
        self.frame.pack(pady=20)

        # Treeview para exibir os dados
        self.tree = ttk.Treeview(self.frame)
        self.tree.pack()

        # Scrollbar
        self.scrollbar = ttk.Scrollbar(self.frame, orient="vertical", command=self.tree.yview)
        self.scrollbar.pack(side='right', fill='y')
        self.tree.configure(yscrollcommand=self.scrollbar.set)

    def load_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx;*.xls")])
        if file_path:
            try:
                df = pd.read_excel(file_path)
                self.show_data(df)
            except Exception as e:
                messagebox.showerror("Erro", f"Não foi possível ler o arquivo:\n{e}")

    def show_data(self, df):
        # Limpa a Treeview antes de mostrar novos dados
        self.tree.delete(*self.tree.get_children())
        
        # Adiciona as colunas
        self.tree["columns"] = list(df.columns)
        self.tree["show"] = "headings"

        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)

        # Adiciona os dados ao Treeview
        for index, row in df.iterrows():
            self.tree.insert("", "end", values=list(row))

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelReaderApp(root)
    root.mainloop()
