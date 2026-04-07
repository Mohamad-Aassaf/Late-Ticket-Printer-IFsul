import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
import win32print
from datetime import datetime
import re
import csv

ALUNOS_FILE = "alunos.json"
ATRASOS_FILE = "atrasos.json"

# ==========================
# GARANTIR ARQUIVOS
# ==========================
def garantir_arquivos():
    if not os.path.exists(ALUNOS_FILE):
        with open(ALUNOS_FILE, "w", encoding="utf-8") as f:
            json.dump({}, f)

    if not os.path.exists(ATRASOS_FILE):
        with open(ATRASOS_FILE, "w", encoding="utf-8") as f:
            json.dump([], f)

garantir_arquivos()

# ==========================
# JSON
# ==========================
def carregar_alunos():
    try:
        with open(ALUNOS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return {}

def salvar_alunos(data):
    with open(ALUNOS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def carregar_atrasos():
    try:
        with open(ATRASOS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except:
        return []

def salvar_atrasos(data):
    with open(ATRASOS_FILE, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

# ==========================
# FORMATAR HORA
# ==========================
def formatar_horario(event):
    texto = re.sub("[^0-9]", "", event.widget.get())
    if len(texto) >= 3:
        texto = texto[:2] + ":" + texto[2:4]
    event.widget.delete(0, tk.END)
    event.widget.insert(0, texto[:5])

# ==========================
# AUTOCOMPLETE INTELIGENTE
# ==========================
def sugerir_aluno(event=None):
    texto_nome = entry_nome.get().strip().lower()
    texto_mat = entry_matricula.get().strip()

    lista_sugestoes.delete(0, tk.END)

    if not texto_nome and not texto_mat:
        return

    alunos = carregar_alunos()
    resultados = []

    for matricula, dados in alunos.items():
        nome = dados["nome"].lower()
        score = 0

        if texto_nome:
            if nome.startswith(texto_nome):
                score += 3
            elif texto_nome in nome:
                score += 1

        if texto_mat:
            if matricula.startswith(texto_mat):
                score += 3
            elif texto_mat in matricula:
                score += 1

        if score > 0:
            resultados.append((score, matricula, dados["nome"]))

    resultados.sort(reverse=True)

    for _, m, n in resultados[:5]:
        lista_sugestoes.insert(tk.END, f"{m} - {n}")

def selecionar_sugestao(event):
    if not lista_sugestoes.curselection():
        return

    selecionado = lista_sugestoes.get(lista_sugestoes.curselection())
    matricula = selecionado.split(" - ")[0]

    alunos = carregar_alunos()
    aluno = alunos.get(matricula, {})

    entry_matricula.delete(0, tk.END)
    entry_matricula.insert(0, matricula)

    entry_nome.delete(0, tk.END)
    entry_nome.insert(0, aluno.get("nome", ""))

    lista_sugestoes.delete(0, tk.END)

# ==========================
# REGISTRAR ATRASO
# ==========================
def registrar_atraso():
    matricula = entry_matricula.get().strip()
    nome = entry_nome.get().strip()

    if not matricula or not nome:
        messagebox.showerror("Erro", "Preencha matrícula e nome.")
        return

    alunos = carregar_alunos()
    alunos[matricula] = {"nome": nome}
    salvar_alunos(alunos)

    registro = {
        "data": datetime.now().strftime("%d/%m/%Y"),
        "hora_registro": datetime.now().strftime("%H:%M"),
        "matricula": matricula,
        "nome": nome,
        "docente": entry_docente.get().strip(),
        "turma": entry_turma.get().strip(),
        "inicio": entry_inicio.get().strip(),
        "chegada": entry_chegada.get().strip(),
        "motivo": combo_motivo.get().strip()
    }

    atrasos = carregar_atrasos()
    atrasos.append(registro)
    salvar_atrasos(atrasos)

    imprimir_termica(registro)

# ==========================
# IMPRESSÃO BONITA
# ==========================
def imprimir_termica(registro):
    printer_name = win32print.GetDefaultPrinter()
    hPrinter = win32print.OpenPrinter(printer_name)

    win32print.StartDocPrinter(hPrinter, 1, ("Atraso", None, "RAW"))
    win32print.StartPagePrinter(hPrinter)

    ESC = b'\x1b'
    GS = b'\x1d'
    conteudo = b''

    conteudo += ESC + b'@'
    conteudo += ESC + b'\x61\x01'
    conteudo += ESC + b'\x21\x20'

    conteudo += "IF - REGISTRO DE ATRASO\n".encode("cp850", errors="replace")

    conteudo += ESC + b'\x21\x00'
    conteudo += ESC + b'\x61\x00'

    conteudo += b"\n================================\n"

    linhas = [
        f"Data: {registro['data']}  ",
        "--------------------------------",
        f"Aluno   : {registro['nome']}",
        f"Matricula: {registro['matricula']}",
        f"Turma   : {registro['turma']}",
        f"Docente : {registro['docente']}",
        "--------------------------------",
        f"Inicio Aula : {registro['inicio']}",
        f"Chegada     : {registro['chegada']}",
        "--------------------------------",
        f"Motivo:",
        f"{registro['motivo']}",
    ]

    for linha in linhas:
        conteudo += (linha + "\n").encode("cp850", errors="replace")

    conteudo += b"\n\n\n"
    conteudo += GS + b'V\x00'
    conteudo += b"\n\n\n\n\n"

    win32print.WritePrinter(hPrinter, conteudo)
    win32print.EndPagePrinter(hPrinter)
    win32print.EndDocPrinter(hPrinter)
    win32print.ClosePrinter(hPrinter)

# ==========================
# VER REGISTROS
# ==========================
def ver_registros():
    atrasos = carregar_atrasos()

    if not atrasos:
        messagebox.showinfo("Aviso", "Nenhum registro encontrado.")
        return

    janela = tk.Toplevel(root)
    janela.title("Registros de Atrasos")
    janela.geometry("1350x720")
    janela.resizable(True, True)
    janela.minsize(800, 600)

    apply_theme_to_window(janela)

    tk.Label(
        janela,
        text="REGISTROS DE ATRASOS",
        font=("Segoe UI", 18, "bold")
    ).pack(pady=10)

    legenda_texto = (
        "DATA | "
        "MATRÍCULA | "
        "NOME | "
        "DOCENTE | "
        "TURMA | "
        "INÍCIO | "
        "CHEGADA | "
        "MOTIVO"
    )

    tk.Label(
        janela,
        text=legenda_texto,
        font=("Segoe UI", 10),
        wraplength=1300,
        justify="left",
        foreground="#333"
    ).pack(padx=20, pady=5)

    frame_busca = tk.Frame(janela)
    frame_busca.pack(pady=10)

    tk.Label(frame_busca, text="Buscar por Matrícula ou Nome:",
             font=("Segoe UI", 11)).pack(side="left", padx=5)

    entry_busca = tk.Entry(frame_busca, font=("Segoe UI", 11), width=35)
    entry_busca.pack(side="left", padx=5)

    lista = tk.Listbox(janela, font=("Courier New", 10), selectmode=tk.SINGLE, exportselection=False)
    lista.pack(fill="both", expand=True, padx=20, pady=10)

    lista_indices = []

    def ordenar_por_aluno(lista_registros):
        return sorted(
            enumerate(lista_registros),
            key=lambda x: (
                x[1].get("matricula", ""),
                x[1].get("data", ""),
                x[1].get("hora_registro", "")
            )
        )

    registros_ordenados = ordenar_por_aluno(atrasos)

    def atualizar_lista(registros):
        lista.delete(0, tk.END)
        lista_indices.clear()

        for original_index, a in registros:
            linha = (
                f"{a.get('data','')} | "
                f"{a.get('matricula','')} | "
                f"{a.get('nome','')} | "
                f"{a.get('docente','')} | "
                f"{a.get('turma','')} | "
                f"Início: {a.get('inicio','')} | "
                f"Chegada: {a.get('chegada','')} | "
                f"{a.get('motivo','')}"
            )
            lista.insert(tk.END, linha)
            lista_indices.append(original_index)

    def procurar():
        texto = entry_busca.get().strip().lower()

        if not texto:
            atualizar_lista(registros_ordenados)
            return

        filtrados = [
            item for item in registros_ordenados
            if texto in item[1].get("matricula", "").lower()
            or texto in item[1].get("nome", "").lower()
        ]

        atualizar_lista(filtrados)

    def exportar_atrasos():
        if not atrasos:
            messagebox.showinfo("Aviso", "Nenhum registro para exportar.")
            return

        arquivo = filedialog.asksaveasfilename(
            title="Exportar registros",
            defaultextension=".csv",
            filetypes=[
                ("CSV", "*.csv"),
                ("Excel XLSX", "*.xlsx"),
                ("Todos", "*.*")
            ],
            initialfile="atrasos.csv"
        )
        if not arquivo:
            return

        ext = os.path.splitext(arquivo)[1].lower()

        if ext == ".xlsx":
            try:
                from openpyxl import Workbook
                wb = Workbook()
                ws = wb.active
                ws.title = "Atrasos"
                ws.append([
                    "Data", "Hora registro", "Matrícula", "Nome",
                    "Docente", "Turma", "Início", "Chegada", "Motivo"
                ])
                for a in atrasos:
                    ws.append([
                        a.get("data", ""),
                        a.get("hora_registro", ""),
                        a.get("matricula", ""),
                        a.get("nome", ""),
                        a.get("docente", ""),
                        a.get("turma", ""),
                        a.get("inicio", ""),
                        a.get("chegada", ""),
                        a.get("motivo", "")
                    ])
                wb.save(arquivo)
                messagebox.showinfo("Exportar", f"Arquivo salvo em:\n{arquivo}")
                return
            except ImportError:
                arquivo = os.path.splitext(arquivo)[0] + ".csv"
                ext = ".csv"
                messagebox.showwarning(
                    "Exportar",
                    "openpyxl não está instalado. O arquivo será salvo como CSV."
                )

        if ext != ".csv":
            arquivo += ".csv"

        with open(arquivo, "w", newline="", encoding="utf-8-sig") as f:
            writer = csv.writer(f, delimiter=";", quoting=csv.QUOTE_MINIMAL)
            writer.writerow([
                "Data", "Hora registro", "Matrícula", "Nome",
                "Docente", "Turma", "Início", "Chegada", "Motivo"
            ])
            for a in atrasos:
                writer.writerow([
                    a.get("data", ""),
                    a.get("hora_registro", ""),
                    a.get("matricula", ""),
                    a.get("nome", ""),
                    a.get("docente", ""),
                    a.get("turma", ""),
                    a.get("inicio", ""),
                    a.get("chegada", ""),
                    a.get("motivo", "")
                ])

        messagebox.showinfo("Exportar", f"Arquivo salvo em:\n{arquivo}")

    atualizar_lista(registros_ordenados)

    tk.Button(
        frame_busca,
        text="Procurar",
        background="#1976d2",
        foreground="white",
        command=procurar
    ).pack(side="left", padx=5)

    tk.Button(
        frame_busca,
        text="Exportar Excel/Sheets",
        background="#4caf50",
        foreground="white",
        command=exportar_atrasos
    ).pack(side="left", padx=5)

    def excluir():
        if not lista.curselection():
            messagebox.showerror("Erro", "Selecione um registro.")
            return

        indice_lista = lista.curselection()[0]
        indice_real = lista_indices[indice_lista]

        if messagebox.askyesno("Confirmar", "Deseja excluir este registro?"):
            atrasos.pop(indice_real)
            salvar_atrasos(atrasos)
            janela.destroy()
            ver_registros()

    def editar():
        if not lista.curselection():
            messagebox.showerror("Erro", "Selecione um registro.")
            return

        indice_lista = lista.curselection()[0]
        indice_real = lista_indices[indice_lista]

        registro = atrasos[indice_real]

        janela_editar = tk.Toplevel(janela)
        janela_editar.title("Editar Registro")
        janela_editar.geometry("500x650")
        janela_editar.resizable(True, True)
        janela_editar.minsize(400, 550)

        apply_theme_to_window(janela_editar)

        def criar_label(txt):
            tk.Label(janela_editar, text=txt,
                     font=("Segoe UI", 11)).pack(anchor="w", pady=(10, 0))

        def criar_entry(valor=""):
            e = tk.Entry(janela_editar, font=("Segoe UI", 11))
            e.pack(fill="x", padx=20)
            e.insert(0, valor)
            return e

        criar_label("Matrícula")
        entry_mat = criar_entry(registro.get("matricula",""))

        criar_label("Nome")
        entry_nome = criar_entry(registro.get("nome",""))

        criar_label("Docente")
        entry_doc = criar_entry(registro.get("docente",""))

        criar_label("Turma")
        entry_turma = criar_entry(registro.get("turma",""))

        criar_label("Início da Aula")
        entry_inicio = criar_entry(registro.get("inicio",""))

        criar_label("Chegada")
        entry_chegada = criar_entry(registro.get("chegada",""))

        criar_label("Motivo")
        entry_motivo = criar_entry(registro.get("motivo",""))

        def salvar_edicao():
            atrasos[indice_real] = {
                "data": registro.get("data",""),
                "hora_registro": registro.get("hora_registro",""),
                "matricula": entry_mat.get().strip(),
                "nome": entry_nome.get().strip(),
                "docente": entry_doc.get().strip(),
                "turma": entry_turma.get().strip(),
                "inicio": entry_inicio.get().strip(),
                "chegada": entry_chegada.get().strip(),
                "motivo": entry_motivo.get().strip()
            }

            salvar_atrasos(atrasos)
            janela_editar.destroy()
            janela.destroy()
            ver_registros()

        tk.Button(
            janela_editar,
            text="Salvar Alterações",
            background="#2e7d32",
            foreground="white",
            command=salvar_edicao
        ).pack(pady=20)

    frame_botoes = tk.Frame(janela)
    frame_botoes.pack(pady=15)

    tk.Button(frame_botoes, text="Editar Registro",
              background="#1976d2", foreground="white",
              width=20, command=editar).pack(side="left", padx=10)

    tk.Button(frame_botoes, text="Excluir Registro",
              background="#c62828", foreground="white",
              width=20, command=excluir).pack(side="left", padx=10)

    tk.Button(frame_botoes, text="Cancelar",
              background="gray", foreground="white",
              width=20, command=janela.destroy).pack(side="left", padx=10)

# ==========================
# ATUALIZAR ALUNO (NOME E MATRICULA)
# ==========================
def atualizar_aluno():
    alunos = carregar_alunos()

    janela = tk.Toplevel(root)
    janela.title("Atualizar Aluno")
    janela.geometry("500x550")
    janela.resizable(True, True)
    janela.minsize(400, 450)

    apply_theme_to_window(janela)

    tk.Label(janela, text="Selecione o aluno:", font=("Segoe UI", 11, "bold")).pack()

    lista = tk.Listbox(janela)
    lista.pack(fill="both", expand=True, pady=10)

    for m, d in alunos.items():
        lista.insert(tk.END, f"{m} - {d['nome']}")

    tk.Label(janela, text="Nova Matrícula").pack()
    entry_mat = tk.Entry(janela)
    entry_mat.pack(fill="x", padx=20, pady=5)

    tk.Label(janela, text="Novo Nome").pack()
    entry_nome_novo = tk.Entry(janela)
    entry_nome_novo.pack(fill="x", padx=20, pady=5)

    def selecionar(event):
        if not lista.curselection():
            return
        selecionado = lista.get(lista.curselection())
        mat = selecionado.split(" - ")[0]
        entry_mat.delete(0, tk.END)
        entry_mat.insert(0, mat)
        entry_nome_novo.delete(0, tk.END)
        entry_nome_novo.insert(0, alunos[mat]["nome"])

    lista.bind("<<ListboxSelect>>", selecionar)

    def salvar():
        if not lista.curselection():
            messagebox.showerror("Erro", "Selecione um aluno.")
            return

        mat_antiga = lista.get(lista.curselection()).split(" - ")[0]
        nova_mat = entry_mat.get().strip()
        novo_nome = entry_nome_novo.get().strip()

        if not nova_mat or not novo_nome:
            messagebox.showerror("Erro", "Preencha os campos.")
            return

        alunos.pop(mat_antiga)
        alunos[nova_mat] = {"nome": novo_nome}
        salvar_alunos(alunos)

        messagebox.showinfo("Sucesso", "Aluno atualizado!")
        janela.destroy()

    tk.Button(janela, text="Salvar", background="#2e7d32", foreground="white", command=salvar).pack(pady=5)
    tk.Button(janela, text="Cancelar", background="gray", foreground="white", command=janela.destroy).pack()

# ==========================
# THEME MANAGEMENT
# ==========================
current_theme = "light"

def apply_theme_to_widget(widget):
    if isinstance(widget, tk.Label):
        if current_theme == "light":
            widget.configure(background="#f5f5f5", foreground="#333")
        else:
            widget.configure(background="#333", foreground="#f5f5f5")
    elif isinstance(widget, tk.Entry):
        if current_theme == "light":
            try:
                widget.configure(background="white", foreground="#333", insertbackground="black")
            except tk.TclError:
                widget.configure(background="white", foreground="#333")
        else:
            try:
                widget.configure(background="#555", foreground="#f5f5f5", insertbackground="#f5f5f5")
            except tk.TclError:
                widget.configure(background="#555", foreground="#f5f5f5")
    elif isinstance(widget, tk.Button):
        # Keep button colors, but ensure bg/fg are set if needed
        pass
    elif isinstance(widget, tk.Listbox):
        if current_theme == "light":
            widget.configure(background="white", foreground="#333")
        else:
            widget.configure(background="#555", foreground="#f5f5f5")
    elif isinstance(widget, ttk.Combobox):
        # For ttk, we can set styles, but for simplicity, leave as is or add custom style
        pass
    elif isinstance(widget, tk.Frame):
        if current_theme == "light":
            widget.configure(background="#f5f5f5")
        else:
            widget.configure(background="#333")
    # Recurse for children
    for child in widget.winfo_children():
        apply_theme_to_widget(child)

def apply_theme():
    if current_theme == "light":
        root.configure(bg="#f5f5f5")
        canvas.configure(bg="#f5f5f5")
    else:
        root.configure(bg="#333")
        canvas.configure(bg="#333")

    apply_theme_to_widget(root)

    if current_theme == "light":
        toggle_button.configure(text="🌙 Dark Mode")
    else:
        toggle_button.configure(text="☀ Light Mode")


def apply_theme_to_window(window):
    apply_theme_to_widget(window)

def toggle_theme():
    global current_theme
    current_theme = "dark" if current_theme == "light" else "light"
    apply_theme()

# ==========================
# INTERFACE
# ==========================
root = tk.Tk()
root.title("Sistema de Atrasos 2026")
root.geometry("820x950")
root.resizable(True, True)
root.minsize(600, 700)

frame = tk.Frame(root, padx=40, pady=40)
frame.pack(fill="both", expand=True)

# Create canvas and scrollbars for responsive design
canvas = tk.Canvas(frame)
scrollbar_y = tk.Scrollbar(frame, orient="vertical", command=canvas.yview)
scrollbar_x = tk.Scrollbar(frame, orient="horizontal", command=canvas.xview)
scrollable_frame = tk.Frame(canvas)

scrollable_frame.bind(
    "<Configure>",
    lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
)

canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="n")
canvas.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)

def update_scrollable_position(event=None):
    canvas_width = canvas.winfo_width()
    canvas.itemconfig(canvas_window, width=canvas_width)

canvas.bind("<Configure>", update_scrollable_position)

# Add mouse wheel scrolling
def _on_mousewheel(event):
    canvas.yview_scroll(int(-1*(event.delta/120)), "units")

def _bind_to_mousewheel(event):
    canvas.bind_all("<MouseWheel>", _on_mousewheel)

def _unbind_from_mousewheel(event):
    canvas.unbind_all("<MouseWheel>")

canvas.bind('<Enter>', _bind_to_mousewheel)
canvas.bind('<Leave>', _unbind_from_mousewheel)

canvas.pack(side="left", fill="both", expand=True)
scrollbar_y.pack(side="right", fill="y")
scrollbar_x.pack(side="bottom", fill="x")

# Dark mode toggle button at the top
toggle_button = tk.Button(scrollable_frame, text="🌙 Dark Mode", command=toggle_theme, font=("Segoe UI", 10))
toggle_button.pack(pady=10)

# Title
title_label = tk.Label(scrollable_frame, text="Registro de Atraso", font=("Segoe UI", 24, "bold"))
title_label.pack(pady=20)

# Container to center the form vertically
container = tk.Frame(scrollable_frame)
container.pack(expand=True)


# Form frame for centering
form_frame = tk.Frame(container)
form_frame.pack(anchor="center")

# Grid layout for fields (2 columns: label and entry)
def create_field(label_text, row, entry_widget=None):
    label = tk.Label(form_frame, text=label_text, font=("Segoe UI", 12))
    label.grid(row=row, column=0, sticky="e", padx=10, pady=5)
    if entry_widget:
        entry_widget.grid(row=row, column=1, sticky="ew", padx=10, pady=5)
    return label

# Matrícula
entry_matricula = tk.Entry(form_frame, font=("Segoe UI", 12))
create_field("Matrícula:", 0, entry_matricula)
entry_matricula.bind("<KeyRelease>", sugerir_aluno)

# Nome
entry_nome = tk.Entry(form_frame, font=("Segoe UI", 12))
create_field("Nome:", 1, entry_nome)
entry_nome.bind("<KeyRelease>", sugerir_aluno)

# Sugestões
lista_sugestoes = tk.Listbox(form_frame, height=4)
lista_sugestoes.grid(row=2, column=0, columnspan=2, sticky="ew", padx=10, pady=5)
lista_sugestoes.bind("<<ListboxSelect>>", selecionar_sugestao)

# Docente
entry_docente = tk.Entry(form_frame, font=("Segoe UI", 12))
create_field("Docente:", 3, entry_docente)

# Turma
entry_turma = tk.Entry(form_frame, font=("Segoe UI", 12))
create_field("Turma:", 4, entry_turma)

# Início
entry_inicio = tk.Entry(form_frame, font=("Segoe UI", 12))
create_field("Início:", 5, entry_inicio)
entry_inicio.bind("<KeyRelease>", formatar_horario)

# Chegada
entry_chegada = tk.Entry(form_frame, font=("Segoe UI", 12))
create_field("Chegada:", 6, entry_chegada)
entry_chegada.bind("<KeyRelease>", formatar_horario)

# Motivo
combo_motivo = ttk.Combobox(form_frame, values=[
    "Transporte privado", "Transporte publico",
    "Casa", "Trabalho", "Medico",
    "Estava no IF", "Acordou tarde", "Transito"
], font=("Segoe UI", 12))
create_field("Motivo:", 7, combo_motivo)

# Buttons
button_frame = tk.Frame(container)
button_frame.pack(anchor="center", pady=20)

tk.Button(button_frame, text="Registrar Atraso",
          background="#2e7d32", foreground="white", font=("Segoe UI", 12),
          command=registrar_atraso).pack(fill="x", pady=5)

tk.Button(button_frame, text="Ver Registros",
          background="#1976d2", foreground="white", font=("Segoe UI", 12),
          command=ver_registros).pack(fill="x", pady=5)

tk.Button(button_frame, text="Atualizar Dados do Aluno",
          background="#ff9800", foreground="white", font=("Segoe UI", 12),
          command=atualizar_aluno).pack(fill="x", pady=5)

# Configure grid weights for centering
form_frame.columnconfigure(1, weight=1)

# Apply initial theme
apply_theme()

root.mainloop()
