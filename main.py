import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
import win32print
from datetime import datetime
import re

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
        f"Data: {registro['data']}  Hora: {registro['hora_registro']}",
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

    # ==============================
    # TÍTULO
    # ==============================
    tk.Label(
        janela,
        text="REGISTROS DE ATRASOS",
        font=("Segoe UI", 18, "bold")
    ).pack(pady=10)

    # ==============================
    # LEGENDA EXPLICATIVA
    # ==============================
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
        fg="#333"
    ).pack(padx=20, pady=5)

    # ==============================
    # CAMPO DE BUSCA
    # ==============================
    frame_busca = tk.Frame(janela)
    frame_busca.pack(pady=10)

    tk.Label(frame_busca, text="Buscar por Matrícula ou Nome:",
             font=("Segoe UI", 11)).pack(side="left", padx=5)

    entry_busca = tk.Entry(frame_busca, font=("Segoe UI", 11), width=35)
    entry_busca.pack(side="left", padx=5)

    lista = tk.Listbox(janela, font=("Courier New", 10))
    lista.pack(fill="both", expand=True, padx=20, pady=10)

    # ==============================
    # ORDENAR POR ALUNO (AGRUPAR)
    # ==============================
    def ordenar_por_aluno(lista_registros):
        return sorted(lista_registros, key=lambda x: (
            x.get("matricula",""),
            x.get("data",""),
            x.get("hora_registro","")
        ))

    registros_ordenados = ordenar_por_aluno(atrasos)

    # ==============================
    # ATUALIZAR LISTA
    # ==============================
    def atualizar_lista(registros):
        lista.delete(0, tk.END)

        for i, a in enumerate(registros):
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

    atualizar_lista(registros_ordenados)

    # ==============================
    # PROCURAR
    # ==============================
    def procurar():
        texto = entry_busca.get().strip().lower()

        if not texto:
            atualizar_lista(registros_ordenados)
            return

        filtrados = []

        for a in registros_ordenados:
            if (
                texto in a.get("matricula","").lower() or
                texto in a.get("nome","").lower()
            ):
                filtrados.append(a)

        atualizar_lista(filtrados)

    tk.Button(
        frame_busca,
        text="Procurar",
        bg="#1976d2",
        fg="white",
        command=procurar
    ).pack(side="left", padx=5)

    # ==============================
    # EXCLUIR
    # ==============================
    def excluir():
        if not lista.curselection():
            messagebox.showerror("Erro", "Selecione um registro.")
            return

        indice_lista = lista.curselection()[0]
        linha = lista.get(indice_lista)
        indice_real = int(linha.split(" | ")[0])

        if messagebox.askyesno("Confirmar", "Deseja excluir este registro?"):
            atrasos.pop(indice_real)
            salvar_atrasos(atrasos)
            janela.destroy()
            ver_registros()

    # ==============================
    # EDITAR
    # ==============================
    def editar():
        if not lista.curselection():
            messagebox.showerror("Erro", "Selecione um registro.")
            return

        indice_lista = lista.curselection()[0]
        linha = lista.get(indice_lista)
        indice_real = int(linha.split(" | ")[0])

        registro = atrasos[indice_real]

        janela_editar = tk.Toplevel(janela)
        janela_editar.title("Editar Registro")
        janela_editar.geometry("500x650")

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
            bg="#2e7d32",
            fg="white",
            command=salvar_edicao
        ).pack(pady=20)

    # ==============================
    # BOTÕES
    # ==============================
    frame_botoes = tk.Frame(janela)
    frame_botoes.pack(pady=15)

    tk.Button(frame_botoes, text="Editar Registro",
              bg="#1976d2", fg="white",
              width=20, command=editar).pack(side="left", padx=10)

    tk.Button(frame_botoes, text="Excluir Registro",
              bg="#c62828", fg="white",
              width=20, command=excluir).pack(side="left", padx=10)

    tk.Button(frame_botoes, text="Cancelar",
              bg="gray", fg="white",
              width=20, command=janela.destroy).pack(side="left", padx=10)

# ==========================
# ATUALIZAR ALUNO (NOME E MATRICULA)
# ==========================
def atualizar_aluno():
    alunos = carregar_alunos()

    janela = tk.Toplevel(root)
    janela.title("Atualizar Aluno")
    janela.geometry("500x550")

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

    tk.Button(janela, text="Salvar", bg="#2e7d32", fg="white", command=salvar).pack(pady=5)
    tk.Button(janela, text="Cancelar", bg="gray", fg="white", command=janela.destroy).pack()

# ==========================
# INTERFACE
# ==========================
root = tk.Tk()
root.title("Sistema de Atrasos 2026")
root.geometry("820x950")

frame = tk.Frame(root, padx=40, pady=40)
frame.pack(fill="both", expand=True)

tk.Label(frame, text="Registro de Atraso", font=("Segoe UI", 22, "bold")).pack(pady=20)

def criar_label(txt):
    tk.Label(frame, text=txt, font=("Segoe UI", 12)).pack(anchor="w", pady=(10, 0))

def criar_entry():
    e = tk.Entry(frame, font=("Segoe UI", 12))
    e.pack(fill="x", ipady=6)
    return e

criar_label("Matrícula")
entry_matricula = criar_entry()
entry_matricula.bind("<KeyRelease>", sugerir_aluno)

criar_label("Nome")
entry_nome = criar_entry()
entry_nome.bind("<KeyRelease>", sugerir_aluno)

lista_sugestoes = tk.Listbox(frame, height=4)
lista_sugestoes.pack(fill="x")
lista_sugestoes.bind("<<ListboxSelect>>", selecionar_sugestao)

criar_label("Docente")
entry_docente = criar_entry()

criar_label("Turma")
entry_turma = criar_entry()

criar_label("Início")
entry_inicio = criar_entry()
entry_inicio.bind("<KeyRelease>", formatar_horario)

criar_label("Chegada")
entry_chegada = criar_entry()
entry_chegada.bind("<KeyRelease>", formatar_horario)

criar_label("Motivo")
combo_motivo = ttk.Combobox(frame, values=[
    "Transporte privado", "Transporte publico",
    "Casa", "Trabalho", "Medico",
    "Estava no IF", "Acordou tarde", "Transito"
])
combo_motivo.pack(fill="x", ipady=5)

tk.Button(frame, text="Registrar Atraso",
          bg="#2e7d32", fg="white",
          command=registrar_atraso).pack(pady=20, fill="x")

tk.Button(frame, text="Ver Registros",
          bg="#1976d2", fg="white",
          command=ver_registros).pack(fill="x", pady=5)

tk.Button(frame, text="Atualizar Dados do Aluno",
          bg="#ff9800", fg="white",
          command=atualizar_aluno).pack(fill="x", pady=5)

root.mainloop()
