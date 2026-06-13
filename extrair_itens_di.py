"""
Extrai Itens DI - YYMMDD.xlsx (v1)
==================================
Autor: Augusto G. Silvestrin (silvestrin) - Estagiario Eng. Produto, DNI - Metalfrio Solutions

EN: Scans all project files in a folder, locates the "RELAÇÃO DE
COMPONENTES" table inside the "Controle de Projeto" sheet of each file,
filters rows where SIT == "DI" (Desenvolvimento de Item), and produces a
consolidated spreadsheet with two sheets:
  - "Itens DI (Todos)": every DI row found, tagged with its source PROJETO
  - "Itens DI (Unicos)": deduplicated by CODIGO (first occurrence wins),
    keeping the source PROJETO as reference, for later cross-checking
    against the Monday "Conteudo - Qualidade" PPAP data.

Requisitos: pip install openpyxl
Build: pyinstaller --onefile --windowed --name ExtrairItensDI --version-file version_info.txt extrair_itens_di.py
"""

__author__ = "Augusto G. Silvestrin (silvestrin)"

import os
import sys
import datetime
import multiprocessing as mp
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
from openpyxl.utils import get_column_letter

# =====================================================================
# CONFIGURACAO
# =====================================================================
PASTA_ORIGEM = r"C:\Projetos"
PASTA_DESTINO = str(Path.home() / "Downloads")
EXTENSOES_VALIDAS = {".xls", ".xlsx", ".xlsm"}

TEXTO_TABELA_COMPONENTES = "RELACAO DE COMPONENTES"
TEXTO_TABELA_PRODUTOS = "CODIGO PRODUTO ACABADO"  # usada so para nao confundir titulos

# Celula onde fica o numero/identificador do projeto (igual ao script de
# Projetos Preliminares)
CELULA_PROJETO = ("Controle de Projeto", "Q15")

# Cabecalhos da tabela "RELACAO DE COMPONENTES" (conforme planilha real)
# nome normalizado -> nome de exibicao na saida
COLUNAS_COMPONENTES = {
    "ITEM": "ITEM",
    "CODIGO": "CODIGO",
    "DESENHO": "DESENHO",
    "LOGIN": "LOGIN",
    "ANT": "ANT",
    "REV": "REV",
    "DENOMINACAO": "DENOMINACAO",
    "TIP": "TIPO",
    "QTDE": "QTDE",
    "UNID": "UNID",
    "ACAO": "ACAO",
    "SIT": "SIT",
    "RIA": "RIA",
    "PLC/RFQ": "PLC/RFQ",
    "APLICACAO": "APLICACAO",
}

# Colunas mantidas na saida (alem de PROJETO), na ordem desejada
COLUNAS_SAIDA = [
    "CODIGO", "DESENHO", "ANT", "REV", "DENOMINACAO",
    "TIPO", "QTDE", "ACAO", "RIA", "PLC/RFQ", "APLICACAO",
]


def _norm(texto):
    """EN: Normalize text for robust matching: uppercase, strip accents,
    collapse internal whitespace."""
    if not isinstance(texto, str):
        return ""
    s = texto.strip().upper()
    substituicoes = {
        "Á": "A", "À": "A", "Â": "A", "Ã": "A",
        "É": "E", "È": "E", "Ê": "E",
        "Í": "I", "Ì": "I",
        "Ó": "O", "Ò": "O", "Ô": "O", "Õ": "O",
        "Ú": "U", "Ù": "U",
        "Ç": "C",
    }
    for k, v in substituicoes.items():
        s = s.replace(k, v)
    return " ".join(s.split())  # colapsa espacos multiplos


def listar_arquivos(pasta):
    arquivos = []
    for entry in os.scandir(pasta):
        if entry.is_file():
            ext = Path(entry.name).suffix.lower()
            if ext in EXTENSOES_VALIDAS and not entry.name.startswith("~$"):
                arquivos.append(entry.path)
    return arquivos


def extrair_itens_di(caminho):
    """
    EN: Locates the "RELAÇÃO DE COMPONENTES" table inside the "Controle de
    Projeto" sheet and returns (itens, diag), where itens is a list of
    dicts {coluna_saida: valor} for rows where SIT == "DI", and diag is a
    short diagnostic string ("ok", "titulo nao encontrado", etc.).
    """
    itens = []

    try:
        wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    except Exception as e:
        return itens, f"erro ao abrir: {e}"

    try:
        if "Controle de Projeto" not in wb.sheetnames:
            return itens, "sem aba 'Controle de Projeto'"
        ws = wb["Controle de Projeto"]

        TEXTO_NORM = _norm(TEXTO_TABELA_COMPONENTES)
        TEXTO_PRODUTOS_NORM = _norm(TEXTO_TABELA_PRODUTOS)

        # Localiza o titulo "RELACAO DE COMPONENTES" (ignora "CODIGO
        # PRODUTO ACABADO", que e' uma tabela diferente na mesma aba)
        linha_titulo = None
        for row_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=min(ws.max_row, 300)), start=1):
            for cell in row:
                texto_cel = _norm(cell.value)
                if TEXTO_NORM in texto_cel and TEXTO_PRODUTOS_NORM not in texto_cel:
                    linha_titulo = row_idx
                    break
            if linha_titulo:
                break

        if linha_titulo is None:
            return itens, "titulo 'RELACAO DE COMPONENTES' nao encontrado"

        # Busca o cabecalho numa janela de algumas linhas abaixo do titulo
        linha_header = None
        mapa_colunas = {}  # nome_normalizado -> col_idx
        for candidata in range(linha_titulo + 1, min(linha_titulo + 6, ws.max_row + 1)):
            mapa_tmp = {}
            for col_idx, cell in enumerate(next(ws.iter_rows(min_row=candidata, max_row=candidata)), start=1):
                texto = _norm(cell.value)
                if texto in COLUNAS_COMPONENTES and texto not in mapa_tmp:
                    mapa_tmp[texto] = col_idx
            # Considera valido se achou pelo menos CODIGO e SIT
            if "CODIGO" in mapa_tmp and "SIT" in mapa_tmp:
                linha_header = candidata
                mapa_colunas = mapa_tmp
                break

        if linha_header is None:
            return itens, "cabecalho da tabela (CODIGO/SIT) nao encontrado"

        col_codigo = mapa_colunas.get("CODIGO")
        col_sit = mapa_colunas.get("SIT")

        # Le todas as linhas de dados ate encontrar CODIGO e ITEM ambos vazios
        # (linha totalmente vazia indica fim da tabela)
        MAX_ITENS_SEGURANCA = 500
        col_item = mapa_colunas.get("ITEM")
        for i in range(MAX_ITENS_SEGURANCA):
            linha_dado = linha_header + 1 + i
            if linha_dado > ws.max_row:
                break

            valor_codigo = ws.cell(row=linha_dado, column=col_codigo).value if col_codigo else None
            valor_item = ws.cell(row=linha_dado, column=col_item).value if col_item else None

            # Linha vazia (sem ITEM nem CODIGO) -> fim da tabela
            if (valor_codigo is None or str(valor_codigo).strip() == "") and \
               (valor_item is None or str(valor_item).strip() == ""):
                break

            valor_sit = ws.cell(row=linha_dado, column=col_sit).value if col_sit else None
            if _norm(valor_sit) != "DI":
                continue

            if valor_codigo is None or str(valor_codigo).strip() == "":
                continue  # SIT=DI mas sem codigo -> ignora

            item = {}
            for nome_norm, nome_saida in COLUNAS_COMPONENTES.items():
                col_idx = mapa_colunas.get(nome_norm)
                item[nome_saida] = ws.cell(row=linha_dado, column=col_idx).value if col_idx else None
            itens.append(item)

        diag = "ok" if itens else f"tabela encontrada (header linha {linha_header}) mas 0 itens com SIT=DI"
    except Exception as e:
        return itens, f"erro: {e}"
    finally:
        wb.close()

    return itens, diag


def ler_arquivo(caminho):
    """
    EN: Worker function (must be top-level/picklable for multiprocessing).
    Returns (caminho, projeto, itens, diag, erro_ou_None).
    """
    projeto = None
    try:
        wb = openpyxl.load_workbook(caminho, data_only=True, read_only=True)
    except Exception as e:
        return (caminho, None, [], None, str(e))

    try:
        aba, celula = CELULA_PROJETO
        if aba in wb.sheetnames:
            try:
                projeto = wb[aba][celula].value
            except Exception:
                projeto = None
    finally:
        wb.close()

    itens, diag = extrair_itens_di(caminho)
    return (caminho, projeto, itens, diag, None)


def main(pasta_origem=None, pasta_destino=None, log=print, progress=None):
    pasta_origem = pasta_origem or PASTA_ORIGEM
    pasta_destino = pasta_destino or PASTA_DESTINO

    def _progress(current, total, message=""):
        if progress:
            progress(current, total, message)

    log("=" * 70)
    log("Extracao de Itens DI - YYMMDD.xlsx (v1)")
    log("=" * 70)

    if not os.path.isdir(pasta_origem):
        log(f"ERRO: pasta de origem nao encontrada: {pasta_origem}")
        return None

    arquivos = listar_arquivos(pasta_origem)
    total = len(arquivos)
    log(f"Encontrados {total} arquivo(s) em {pasta_origem}\n")
    _progress(0, total, "Iniciando...")

    if total == 0:
        log("Nenhum arquivo para processar. Encerrando.")
        return None

    inicio = datetime.datetime.now()

    n_workers = max(1, min(os.cpu_count() or 4, 8))
    log(f"Lendo arquivos em paralelo usando {n_workers} processo(s)...\n")

    # Tabela 1: todos os itens DI, com PROJETO de origem
    todos = []  # list of dicts: {"PROJETO": ..., **item}
    erros = []
    for i, (caminho, projeto, itens, diag, erro) in enumerate(
            _pool_imap(n_workers, ler_arquivo, arquivos), start=1):
        nome = os.path.basename(caminho)
        if erro:
            erros.append((nome, erro))
            log(f"[{i}/{total}] ERRO: {nome} -> {erro}")
        else:
            log(f"[{i}/{total}] OK: {nome} -> {len(itens)} item(ns) DI"
                f"{'' if diag == 'ok' or itens else f' ({diag})'}")
            for item in itens:
                linha = {"PROJETO": projeto}
                linha.update(item)
                todos.append(linha)
        _progress(i, total, f"Lendo arquivos... ({i}/{total})")

    tempo_leitura = (datetime.datetime.now() - inicio).total_seconds()
    log(f"\nLeitura concluida em {tempo_leitura:.1f}s "
        f"({total - len(erros)} ok, {len(erros)} com erro)")
    log(f"Total de itens DI encontrados (todos os projetos): {len(todos)}")

    # Tabela 2: deduplicado por CODIGO (1a ocorrencia)
    unicos = []
    vistos = set()
    for linha in todos:
        codigo = linha.get("CODIGO")
        chave = str(codigo).strip() if codigo is not None else None
        if chave is None or chave == "" or chave in vistos:
            continue
        vistos.add(chave)
        unicos.append(linha)

    log(f"Total de itens DI unicos (por CODIGO): {len(unicos)}")

    # --- Monta o workbook de saida ---
    wb_out = openpyxl.Workbook()

    cabecalhos = ["PROJETO"] + COLUNAS_SAIDA

    _escrever_aba(wb_out.active, "Itens DI (Todos)", cabecalhos, todos)
    ws2 = wb_out.create_sheet("Itens DI (Unicos)")
    _escrever_aba(ws2, None, cabecalhos, unicos, ja_existe=True)

    hoje = datetime.date.today().strftime("%y%m%d")
    nome_saida = f"Itens DI - {hoje}.xlsx"
    caminho_saida = os.path.join(pasta_destino, nome_saida)
    wb_out.save(caminho_saida)

    tempo_total = (datetime.datetime.now() - inicio).total_seconds()
    log(f"\nArquivo salvo em: {caminho_saida}")
    log(f"Tempo total: {tempo_total:.1f}s")
    if erros:
        log(f"\n{len(erros)} arquivo(s) com erro (pulados):")
        for nome, msg in erros:
            log(f"  - {nome}: {msg}")
    log("Concluido.")
    _progress(total, total, "Concluido")
    return caminho_saida


def _pool_imap(n_workers, func, arquivos):
    """EN: Thin wrapper so main() doesn't need to manage Pool lifecycle
    inline (keeps the structure close to the Projetos Preliminares
    script)."""
    with mp.Pool(processes=n_workers) as pool:
        for resultado in pool.imap_unordered(func, arquivos):
            yield resultado


def _escrever_aba(ws, titulo_aba, cabecalhos, linhas, ja_existe=False):
    """EN: Writes header + data rows into ws, applies basic formatting
    (header style, borders, autofilter, frozen header, column widths)."""
    if titulo_aba:
        ws.title = titulo_aba

    fonte_header = Font(bold=True, color="404040")
    fill_header = PatternFill(start_color="D9D9D9", end_color="D9D9D9", fill_type="solid")
    fina = Side(style="thin", color="BFBFBF")
    borda = Border(left=fina, right=fina, top=fina, bottom=fina)

    for j, titulo in enumerate(cabecalhos, start=1):
        c = ws.cell(row=1, column=j, value=titulo)
        c.font = fonte_header
        c.fill = fill_header
        c.alignment = Alignment(horizontal="center", vertical="center")
        c.border = borda

    for i, linha in enumerate(linhas, start=2):
        for j, cab in enumerate(cabecalhos, start=1):
            c = ws.cell(row=i, column=j, value=linha.get(cab))
            c.border = borda
            c.alignment = Alignment(horizontal="center", vertical="center")

    ultima_linha = len(linhas) + 1
    total_colunas = len(cabecalhos)
    ws.auto_filter.ref = f"A1:{get_column_letter(total_colunas)}{ultima_linha}"
    ws.freeze_panes = "A2"

    for col, titulo in enumerate(cabecalhos, start=1):
        largura = max(10, len(str(titulo)) + 2)
        ws.column_dimensions[get_column_letter(col)].width = largura


def executar_gui():
    """
    EN: tkinter GUI: editable source/destination folders, a "Gerar" button,
    a determinate progress bar with a short status line, an optional
    collapsible "Detalhes" log area (hidden by default), and a completion
    popup with the output file path. Processing runs in a background
    thread so the window stays responsive.
    """
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    import threading

    root = tk.Tk()
    root.title("Extrair Itens DI - DNI")
    root.geometry("560x260")
    root.resizable(True, True)

    # --- Linha: Pasta de origem ---
    frame_origem = tk.Frame(root)
    frame_origem.pack(fill="x", padx=12, pady=(14, 5))
    tk.Label(frame_origem, text="Pasta de origem (C:\\Projetos):", width=24, anchor="w").pack(side="left")
    var_origem = tk.StringVar(value=PASTA_ORIGEM)
    entry_origem = tk.Entry(frame_origem, textvariable=var_origem)
    entry_origem.pack(side="left", fill="x", expand=True, padx=5)

    def escolher_origem():
        d = filedialog.askdirectory(initialdir=var_origem.get() or "/")
        if d:
            var_origem.set(d)

    tk.Button(frame_origem, text="...", command=escolher_origem, width=3).pack(side="left")

    # --- Linha: Pasta de destino ---
    frame_destino = tk.Frame(root)
    frame_destino.pack(fill="x", padx=12, pady=5)
    tk.Label(frame_destino, text="Pasta de destino (Downloads):", width=24, anchor="w").pack(side="left")
    var_destino = tk.StringVar(value=PASTA_DESTINO)
    entry_destino = tk.Entry(frame_destino, textvariable=var_destino)
    entry_destino.pack(side="left", fill="x", expand=True, padx=5)

    def escolher_destino():
        d = filedialog.askdirectory(initialdir=var_destino.get() or "/")
        if d:
            var_destino.set(d)

    tk.Button(frame_destino, text="...", command=escolher_destino, width=3).pack(side="left")

    # --- Botao Gerar + status ---
    frame_botao = tk.Frame(root)
    frame_botao.pack(fill="x", padx=12, pady=(15, 5))
    btn_gerar = tk.Button(frame_botao, text="Gerar", width=14)
    btn_gerar.pack(side="left")
    lbl_status = tk.Label(frame_botao, text="Pronto.", anchor="w")
    lbl_status.pack(side="left", padx=10, fill="x", expand=True)

    # --- Barra de progresso ---
    frame_progress = tk.Frame(root)
    frame_progress.pack(fill="x", padx=12, pady=(0, 8))
    pbar = ttk.Progressbar(frame_progress, orient="horizontal", mode="determinate")
    pbar.pack(fill="x")

    # --- Detalhes (log) - colapsavel, oculto por padrao ---
    frame_detalhes_container = tk.Frame(root)
    frame_detalhes_container.pack(fill="both", expand=True, padx=12, pady=(0, 5))

    log_widget = None
    detalhes_visivel = tk.BooleanVar(value=False)

    def alternar_detalhes():
        nonlocal log_widget
        if detalhes_visivel.get():
            if log_widget is None:
                log_widget = scrolledtext.ScrolledText(
                    frame_detalhes_container, state="disabled",
                    font=("Consolas", 8), height=10
                )
            log_widget.pack(fill="both", expand=True, pady=(5, 0))
            btn_detalhes.config(text="▲ Ocultar detalhes")
            root.geometry("560x460")
        else:
            if log_widget is not None:
                log_widget.pack_forget()
            btn_detalhes.config(text="▼ Mostrar detalhes")
            root.geometry("560x260")

    def toggle_detalhes():
        detalhes_visivel.set(not detalhes_visivel.get())
        alternar_detalhes()

    btn_detalhes = tk.Button(root, text="▼ Mostrar detalhes", relief="flat",
                              fg="gray", command=toggle_detalhes)
    btn_detalhes.pack(side="bottom", pady=(0, 2))

    def log(msg):
        def _append():
            if log_widget is not None:
                log_widget.configure(state="normal")
                log_widget.insert("end", str(msg) + "\n")
                log_widget.see("end")
                log_widget.configure(state="disabled")
        root.after(0, _append)

    def progress(current, total, message=""):
        def _update():
            if total > 0:
                pbar.configure(maximum=total, value=current)
            if message:
                lbl_status.config(text=message)
        root.after(0, _update)

    def rodar():
        btn_gerar.config(state="disabled")
        lbl_status.config(text="Iniciando...")
        pbar.configure(value=0)
        if log_widget is not None:
            log_widget.configure(state="normal")
            log_widget.delete("1.0", "end")
            log_widget.configure(state="disabled")

        def trabalho():
            try:
                caminho_saida = main(
                    pasta_origem=var_origem.get(),
                    pasta_destino=var_destino.get(),
                    log=log,
                    progress=progress,
                )
            except Exception as e:
                log(f"\nERRO GERAL: {e}")
                caminho_saida = None

            def finalizar():
                btn_gerar.config(state="normal")
                lbl_status.config(text="Pronto." if caminho_saida else "Falhou.")
                if caminho_saida:
                    messagebox.showinfo(
                        "Concluído",
                        f"Arquivo gerado com sucesso:\n\n{caminho_saida}"
                    )
                else:
                    messagebox.showwarning(
                        "Atenção",
                        "Processamento finalizado sem gerar arquivo.\n"
                        "Clique em 'Mostrar detalhes' e gere novamente para ver o log."
                    )

            root.after(0, finalizar)

        threading.Thread(target=trabalho, daemon=True).start()

    btn_gerar.config(command=rodar)

    # --- Rodape: autor ---
    tk.Label(root, text=f"Autor: {__author__} | DNI - Metalfrio Solutions",
             font=("Segoe UI", 8), fg="gray").pack(side="bottom", pady=(0, 5))

    root.mainloop()


if __name__ == "__main__":
    mp.freeze_support()
    executar_gui()
