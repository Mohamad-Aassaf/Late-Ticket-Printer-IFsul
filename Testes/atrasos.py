import tkinter as tk
from tkinter import ttk, messagebox
import win32print

# ==============================
# CONFIGURAÇÃO DA IMPRESSORA
# ==============================

PRINTER_NAME = "GP-C80250 Series"
def imprimir(texto):
    try:
        hprinter = win32print.OpenPrinter(PRINTER_NAME)
        hjob = win32print.StartDocPrinter(hprinter, 1, ("Cupom", None, "RAW"))
        win32print.StartPagePrinter(hprinter)

        CUT = b'\x1d\x56\x00'  # cortar papel
        win32print.WritePrinter(hprinter, texto.encode("cp850") + CUT)

        win32print.EndPagePrinter(hprinter)
        win32print.EndDocPrinter(hprinter)
        win32print.ClosePrinter(hprinter)

        messagebox.showinfo("Sucesso", "Impresso com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", str(e))


# ==============================
# INTERFACE
# ==============================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Sistema POS - 80mm")
        self.geometry("500x500")
        self.configure(bg="#f4f4f4")

        self.produtos = []

        self.criar_interface()

    def criar_interface(self):
        style = ttk.Style()
        style.configure("TButton", font=("Segoe UI", 10))
        style.configure("TLabel", font=("Segoe UI", 10))

        ttk.Label(self, text="Produto:").pack(pady=5)
        self.entry_nome = ttk.Entry(self, width=40)
        self.entry_nome.pack()

        ttk.Label(self, text="Preço (R$):").pack(pady=5)
        self.entry_preco = ttk.Entry(self, width=20)
        self.entry_preco.pack()

        ttk.Button(self, text="Adicionar Produto", command=self.adicionar_produto).pack(pady=10)

        self.lista = tk.Listbox(self, width=60, height=10)
        self.lista.pack(pady=10)

        self.label_total = ttk.Label(self, text="TOTAL: R$ 0.00", font=("Segoe UI", 12, "bold"))
        self.label_total.pack(pady=10)

        ttk.Button(self, text="Imprimir Cupom", command=self.gerar_cupom).pack(pady=20)

    def adicionar_produto(self):
        nome = self.entry_nome.get()
        preco = self.entry_preco.get()

        if not nome or not preco:
            return

        try:
            preco = float(preco)
        except:
            messagebox.showerror("Erro", "Preço inválido")
            return

        self.produtos.append((nome, preco))
        self.lista.insert(tk.END, f"{nome} - R$ {preco:.2f}")
        self.atualizar_total()

        self.entry_nome.delete(0, tk.END)
        self.entry_preco.delete(0, tk.END)

    def atualizar_total(self):
        total = sum(p[1] for p in self.produtos)
        self.label_total.config(text=f"TOTAL: R$ {total:.2f}")

    def gerar_cupom(self):
        if not self.produtos:
            return

        texto = ""
        texto += "MINHA LOJA\n"
        texto += "Rua Exemplo 123\n"
        texto += "-" * 32 + "\n"

        total = 0
        for nome, preco in self.produtos:
            linha = f"{nome[:20]:20} R$ {preco:6.2f}\n"
            texto += linha
            total += preco

        texto += "-" * 32 + "\n"
        texto += f"TOTAL: R$ {total:.2f}\n\n"
        texto += "Obrigado pela preferencia!\n\n\n"

        imprimir(texto)


if __name__ == "__main__":
    app = App()
    app.mainloop()
