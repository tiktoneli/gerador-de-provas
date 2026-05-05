import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from gerador import (
    carregar_banco,
    carregar_config,
    _calcular_cotas,
    distribuicao_efetiva,
    gerar_prova,
    validar_banco,
)

BASE_DIR = Path(__file__).parent

try:
    config = carregar_config()
except Exception as e:
    import sys

    messagebox.showerror("Erro de configuração", str(e))
    sys.exit(1)

NOMES_NIVEL = {
    1: "Muito fácil",
    2: "Fácil",
    3: "Médio",
    4: "Difícil",
    5: "Muito difícil",
}

assuntos_disponiveis: dict[str, list[str]] = {}


def log(msg: str) -> None:
    area_log.configure(state="normal")
    area_log.insert(tk.END, msg + "\n")
    area_log.see(tk.END)
    area_log.configure(state="disabled")


def escolher_csv() -> None:
    caminho = filedialog.askopenfilename(
        title="Selecionar banco de questões",
        filetypes=[("CSV", "*.csv")],
        initialdir=BASE_DIR,
    )
    if caminho:
        entrada_csv.set(caminho)


def _ler_distribuicao_gui() -> dict[int, float] | None:
    modelo_nome = combo_banca.get()
    dif_range = config[modelo_nome].get("dificuldade_range", [1, 5])
    total_questoes = _total_questoes_modelo(modelo_nome)
    contagens: dict[int, int] = {}
    for nivel in range(dif_range[0], dif_range[1] + 1):
        txt = dist_vars[nivel].get().strip()
        try:
            contagens[nivel] = int(txt)
        except ValueError:
            return None
        if contagens[nivel] < 0:
            return None
    if sum(contagens.values()) != total_questoes:
        return None
    return {nivel: qtd / total_questoes for nivel, qtd in contagens.items()}


def _total_questoes_modelo(modelo_nome: str) -> int:
    total = 0
    for dia in config[modelo_nome]["dias"].values():
        for bloco in dia["blocos"]:
            total += int(bloco["quantidade"])
    return total


def _soma_distribuicao_informada() -> int:
    modelo_nome = combo_banca.get()
    dif_range = config[modelo_nome].get("dificuldade_range", [1, 5])
    soma = 0
    for nivel in range(dif_range[0], dif_range[1] + 1):
        txt = dist_vars[nivel].get().strip()
        try:
            valor = int(txt)
        except ValueError:
            continue
        soma += max(0, valor)
    return soma


def _atualizar_contador_distribuicao(event=None) -> None:
    modelo_nome = combo_banca.get()
    total = _total_questoes_modelo(modelo_nome)
    soma = _soma_distribuicao_informada()
    if soma < total:
        lbl_contador_dist.configure(
            text=f"Faltam {total - soma} questão(ões) para atingir {total}.",
            foreground="#8a6d00",
        )
    elif soma > total:
        lbl_contador_dist.configure(
            text=f"Excedeu em {soma - total} questão(ões). Alvo: {total}.",
            foreground="#8a1f11",
        )
    else:
        lbl_contador_dist.configure(
            text=f"Distribuição completa: {soma}/{total} questões.",
            foreground="#1f6b2a",
        )


def _compativel_modelo(modelos_txt: str, modelo_nome: str) -> bool:
    modelos_txt = (modelos_txt or "").strip()
    if not modelos_txt:
        return True
    return modelo_nome in [x.strip().lower() for x in modelos_txt.split(";")]


def _disciplinas_modelo(modelo_nome: str) -> list[str]:
    disciplinas: list[str] = []
    for dia in config[modelo_nome]["dias"].values():
        for bloco in dia["blocos"]:
            disc = bloco["disciplina"]
            if disc not in disciplinas:
                disciplinas.append(disc)
    return disciplinas


def _set_assuntos_disciplina(event=None) -> None:
    disc = combo_disciplina_assunto.get().strip()
    combo_assunto_assunto["values"] = assuntos_disponiveis.get(disc, [])
    if combo_assunto_assunto["values"]:
        combo_assunto_assunto.set(combo_assunto_assunto["values"][0])
    else:
        combo_assunto_assunto.set("")


def _limpar_tabela_assuntos() -> None:
    for iid in tree_assuntos.get_children():
        tree_assuntos.delete(iid)


def mapear_assuntos() -> None:
    csv_path = entrada_csv.get().strip()
    if not csv_path:
        messagebox.showwarning("Aviso", "Selecione o banco de questões (CSV) primeiro.")
        return

    try:
        df = carregar_banco(csv_path)
        modelo_nome = combo_banca.get()
        disciplinas = _disciplinas_modelo(modelo_nome)

        assuntos_disponiveis.clear()
        for disc in disciplinas:
            mask_disc = df["disciplina"].fillna("").str.strip().str.lower().eq(disc.strip().lower())
            mask_mod = df["modelos"].fillna("").apply(lambda m: _compativel_modelo(str(m), modelo_nome))
            assuntos = sorted(
                {
                    str(a).strip()
                    for a in df[mask_disc & mask_mod]["assunto"].fillna("")
                    if str(a).strip()
                }
            )
            assuntos_disponiveis[disc] = assuntos

        combo_disciplina_assunto["values"] = disciplinas
        if disciplinas:
            combo_disciplina_assunto.set(disciplinas[0])
            _set_assuntos_disciplina()

        _limpar_tabela_assuntos()
        total_assuntos = sum(len(v) for v in assuntos_disponiveis.values())
        log(
            f"Mapa carregado: {len(disciplinas)} disciplina(s), {total_assuntos} assunto(s). "
            "Agora escolha disciplina/assunto e defina o mínimo."
        )
    except Exception as e:
        log(f"✗ Erro ao mapear assuntos: {e}")


def adicionar_ou_atualizar_regra() -> None:
    disc = combo_disciplina_assunto.get().strip()
    assunto = combo_assunto_assunto.get().strip()
    minimo_txt = entrada_minimo_assunto.get().strip()

    if not disc or not assunto:
        messagebox.showwarning("Aviso", "Selecione disciplina e assunto.")
        return

    try:
        minimo = int(minimo_txt)
    except ValueError:
        messagebox.showwarning("Aviso", "Mínimo deve ser um número inteiro.")
        return

    if minimo < 0:
        messagebox.showwarning("Aviso", "Mínimo não pode ser negativo.")
        return

    existente = None
    for iid in tree_assuntos.get_children():
        vals = tree_assuntos.item(iid, "values")
        if vals and vals[0] == disc and vals[1] == assunto:
            existente = iid
            break

    if minimo == 0:
        if existente is not None:
            tree_assuntos.delete(existente)
        return

    if existente is not None:
        tree_assuntos.item(existente, values=(disc, assunto, str(minimo)))
    else:
        tree_assuntos.insert("", tk.END, values=(disc, assunto, str(minimo)))


def remover_regra_selecionada() -> None:
    sel = tree_assuntos.selection()
    if not sel:
        messagebox.showwarning("Aviso", "Selecione uma regra na tabela.")
        return
    for iid in sel:
        tree_assuntos.delete(iid)


def limpar_regras() -> None:
    _limpar_tabela_assuntos()


def _selecionar_regra_tabela(event=None) -> None:
    sel = tree_assuntos.selection()
    if not sel:
        return
    vals = tree_assuntos.item(sel[0], "values")
    if not vals:
        return
    disc, assunto, minimo = vals
    combo_disciplina_assunto.set(disc)
    _set_assuntos_disciplina()
    combo_assunto_assunto.set(assunto)
    entrada_minimo_assunto.delete(0, tk.END)
    entrada_minimo_assunto.insert(0, minimo)


def _ler_regras_assunto_gui() -> dict[str, dict[str, int]]:
    regras: dict[str, dict[str, int]] = {}
    for iid in tree_assuntos.get_children():
        disc, assunto, minimo_txt = tree_assuntos.item(iid, "values")
        minimo = int(minimo_txt)
        if minimo > 0:
            regras.setdefault(str(disc), {})[str(assunto)] = minimo
    return regras


def _limpar_validacao_visual() -> None:
    for iid in tree_validacao.get_children():
        tree_validacao.delete(iid)
    lbl_resumo_validacao.configure(text="Sem validação executada.")


def _add_validacao(tipo: str, status: str, detalhe: str) -> None:
    status_norm = status.strip().lower()
    tag = ""
    if status_norm == "erro":
        tag = "erro"
    elif status_norm == "aviso":
        tag = "aviso"
    elif status_norm == "ok":
        tag = "ok"
    tree_validacao.insert("", tk.END, values=(tipo, status, detalhe), tags=(tag,))


def atualizar_painel_distribuicao(event=None) -> None:
    modelo_nome = combo_banca.get()
    modelo = config[modelo_nome]
    dif_range = modelo.get("dificuldade_range", [1, 5])
    dist = distribuicao_efetiva(modelo)
    total_questoes = _total_questoes_modelo(modelo_nome)
    cotas = _calcular_cotas(total_questoes, dist)

    lbl_range.configure(
        text=(
            f"Faixa válida: {dif_range[0]} - {dif_range[1]} | "
            f"Total da banca: {total_questoes} questões"
        )
    )

    for nivel in range(1, 6):
        in_range = dif_range[0] <= nivel <= dif_range[1]
        if in_range:
            dist_vars[nivel].set(str(cotas.get(nivel, 0)))
            dist_labels[nivel].grid()
            dist_entries[nivel].grid()
            dist_qtd_labels[nivel].grid()
        else:
            dist_labels[nivel].grid_remove()
            dist_entries[nivel].grid_remove()
            dist_qtd_labels[nivel].grid_remove()

    btn_reset_dist.configure(command=lambda: atualizar_painel_distribuicao())
    _atualizar_contador_distribuicao()


def executar_validacao() -> None:
    _limpar_validacao_visual()

    csv_path = entrada_csv.get().strip()
    if not csv_path:
        messagebox.showwarning("Aviso", "Selecione o banco de questões (CSV) primeiro.")
        _add_validacao("Entrada", "Erro", "CSV não selecionado.")
        lbl_resumo_validacao.configure(text="Validação interrompida: selecione um CSV.")
        return

    aplicar_distribuicao = not ignorar_dificuldade_var.get()
    dist = None
    if aplicar_distribuicao:
        dist = _ler_distribuicao_gui()
        if dist is None:
            messagebox.showwarning(
                "Aviso",
                "Verifique as quantidades da distribuição de dificuldades.",
            )
            _add_validacao("Entrada", "Erro", "Distribuição de dificuldades inválida.")
            lbl_resumo_validacao.configure(text="Validação interrompida: distribuição inválida.")
            return

    regras_assunto = {} if ignorar_assuntos_var.get() else _ler_regras_assunto_gui()

    try:
        df = carregar_banco(csv_path)
        modelo_nome = combo_banca.get()
        _add_validacao("Contexto", "Info", f"Banca selecionada: {modelo_nome.upper()}")
        _add_validacao("Contexto", "Info", f"Total de questões no CSV: {df.shape[0]}")
        resultado = validar_banco(
            df,
            modelo_nome,
            config,
            dist,
            regras_assunto,
            aplicar_distribuicao,
        )
        erros = resultado["erros"]
        avisos = resultado["avisos"]

        for e in erros:
            _add_validacao("Disponibilidade", "Erro", e)
        for a in avisos:
            _add_validacao("Distribuição", "Aviso", a)

        if erros:
            log("ERRO - questões insuficientes (bloqueia geração):")
            for e in erros:
                log(f"  {e}")
        if avisos:
            log("Aviso - distribuição aproximada (geração prossegue):")
            for a in avisos:
                log(f"  {a}")
        if not erros and not avisos:
            total = df.shape[0]
            log(f"Banco válido para {modelo_nome.upper()} - {total} questões carregadas.")
            _add_validacao("Resultado", "OK", "Sem erros e sem avisos.")

        lbl_resumo_validacao.configure(
            text=(
                f"Validação da banca {modelo_nome.upper()}: "
                f"{len(erros)} erro(s), {len(avisos)} aviso(s)."
            )
        )
    except Exception as e:
        log(f"✗ Erro na validação: {e}")
        _add_validacao("Execução", "Erro", str(e))
        lbl_resumo_validacao.configure(text="Validação falhou por erro de execução.")


def _thread_gerar(
    csv_path: str,
    modelo_nome: str,
    semente: int | None,
    dist: dict[int, float],
    regras_assunto: dict[str, dict[str, int]],
    aplicar_distribuicao: bool,
) -> None:
    log(
        f"\n-> Gerando {modelo_nome.upper()} | "
        f"semente: {semente if semente is not None else 'aleatória'}"
    )
    try:
        arquivos = gerar_prova(
            modelo_nome,
            csv_path,
            semente,
            dist,
            regras_assunto,
            aplicar_distribuicao,
        )
        pasta = arquivos[0].parent.name
        log(f"✓ Pasta: {pasta}")
        for arq in arquivos:
            log(f"  {arq.name}")
        nomes = "\n".join(a.name for a in arquivos)
        app.after(
            0,
            lambda: messagebox.showinfo("Pronto", f"Prova gerada!\nPasta: {pasta}\n\n{nomes}"),
        )
    except Exception as e:
        msg = str(e)
        log(f"✗ Erro: {msg}")
        app.after(0, lambda m=msg: messagebox.showerror("Erro", m))
    finally:
        app.after(0, lambda: btn_gerar.configure(state="normal"))


def executar_geracao() -> None:
    csv_path = entrada_csv.get().strip()
    if not csv_path:
        messagebox.showwarning("Aviso", "Selecione o banco de questões (CSV) primeiro.")
        return

    aplicar_distribuicao = not ignorar_dificuldade_var.get()
    dist = None
    if aplicar_distribuicao:
        dist = _ler_distribuicao_gui()
        if dist is None:
            messagebox.showwarning(
                "Aviso",
                "Verifique as quantidades da distribuição de dificuldades.",
            )
            return

    regras_assunto = {} if ignorar_assuntos_var.get() else _ler_regras_assunto_gui()

    semente_txt = entrada_semente.get().strip()
    semente = int(semente_txt) if semente_txt.isdigit() else None
    modelo_nome = combo_banca.get()

    btn_gerar.configure(state="disabled")
    threading.Thread(
        target=_thread_gerar,
        args=(csv_path, modelo_nome, semente, dist, regras_assunto, aplicar_distribuicao),
        daemon=True,
    ).start()


app = tk.Tk()
app.title("Gerador de Provas")
app.resizable(True, True)
app.minsize(920, 700)

# Container com scroll vertical
container = ttk.Frame(app)
container.pack(fill="both", expand=True)

canvas = tk.Canvas(container, highlightthickness=0)
scrollbar_y = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
canvas.configure(yscrollcommand=scrollbar_y.set)

scrollbar_y.pack(side="right", fill="y")
canvas.pack(side="left", fill="both", expand=True)

frame = ttk.Frame(canvas, padding=16)
canvas_window = canvas.create_window((0, 0), window=frame, anchor="nw")


def _on_frame_configure(event=None) -> None:
    canvas.configure(scrollregion=canvas.bbox("all"))


def _on_canvas_configure(event) -> None:
    canvas.itemconfigure(canvas_window, width=event.width)


def _on_mousewheel(event) -> None:
    # Delta vem em múltiplos de 120 no Windows.
    canvas.yview_scroll(int(-event.delta / 120), "units")


frame.bind("<Configure>", _on_frame_configure)
canvas.bind("<Configure>", _on_canvas_configure)
canvas.bind_all("<MouseWheel>", _on_mousewheel)

# CSV
ttk.Label(frame, text="Banco de questões:").grid(row=0, column=0, sticky="w")
entrada_csv = tk.StringVar(value=str(BASE_DIR / "banco_questoes.csv"))
ttk.Entry(frame, textvariable=entrada_csv, width=44).grid(row=0, column=1, padx=4)
ttk.Button(frame, text="...", width=3, command=escolher_csv).grid(row=0, column=2)

# Banca
ttk.Label(frame, text="Banca:").grid(row=1, column=0, sticky="w", pady=8)
combo_banca = ttk.Combobox(frame, values=list(config.keys()), state="readonly", width=22)
combo_banca.set(list(config.keys())[0])
combo_banca.grid(row=1, column=1, sticky="w", padx=4)
combo_banca.bind("<<ComboboxSelected>>", atualizar_painel_distribuicao)

# Semente
ttk.Label(frame, text="Semente aleatória:").grid(row=2, column=0, sticky="w")
entrada_semente = ttk.Entry(frame, width=14)
entrada_semente.grid(row=2, column=1, sticky="w", padx=4)
ttk.Label(frame, text="(vazio = aleatório)").grid(row=2, column=2, sticky="w")

# Distribuição de dificuldades
frame_dist = ttk.LabelFrame(frame, text="Distribuição de dificuldades (quantidade)", padding=8)
frame_dist.grid(row=3, column=0, columnspan=3, sticky="ew", pady=6)

lbl_range = ttk.Label(frame_dist, text="")
lbl_range.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 4))

dist_vars: dict[int, tk.StringVar] = {}
dist_labels: dict[int, ttk.Label] = {}
dist_entries: dict[int, ttk.Entry] = {}
dist_qtd_labels: dict[int, ttk.Label] = {}

for i, nivel in enumerate(range(1, 6)):
    dist_vars[nivel] = tk.StringVar()
    dist_vars[nivel].trace_add("write", _atualizar_contador_distribuicao)

    lbl = ttk.Label(frame_dist, text=f"Nível {nivel}  {NOMES_NIVEL[nivel]}:")
    lbl.grid(row=i + 1, column=0, sticky="w", pady=1, padx=(0, 6))
    dist_labels[nivel] = lbl

    ent = ttk.Entry(frame_dist, textvariable=dist_vars[nivel], width=8)
    ent.grid(row=i + 1, column=1, sticky="w")
    dist_entries[nivel] = ent

    qtd_lbl = ttk.Label(frame_dist, text="questões")
    qtd_lbl.grid(row=i + 1, column=2, sticky="w", padx=(2, 0))
    dist_qtd_labels[nivel] = qtd_lbl

ignorar_dificuldade_var = tk.BooleanVar(value=False)
ttk.Checkbutton(
    frame_dist,
    text="Relevar distribuição de dificuldades",
    variable=ignorar_dificuldade_var,
).grid(row=6, column=0, columnspan=3, sticky="w", pady=(6, 0))

btn_reset_dist = ttk.Button(frame_dist, text="Restaurar padrão", width=16)
btn_reset_dist.grid(row=7, column=0, columnspan=3, pady=(8, 0), sticky="w")

lbl_contador_dist = ttk.Label(frame_dist, text="")
lbl_contador_dist.grid(row=8, column=0, columnspan=3, sticky="w", pady=(6, 0))

ttk.Separator(frame, orient="horizontal").grid(
    row=4, column=0, columnspan=3, sticky="ew", pady=10
)

# Assuntos por disciplina (mínimos)
frame_assuntos = ttk.LabelFrame(frame, text="Assuntos por disciplina (mínimos)", padding=8)
frame_assuntos.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 8))

ttk.Button(frame_assuntos, text="1) Mapear assuntos do banco", command=mapear_assuntos).grid(
    row=0, column=0, columnspan=6, sticky="w", pady=(0, 6)
)

ttk.Label(frame_assuntos, text="Disciplina:").grid(row=1, column=0, sticky="w")
combo_disciplina_assunto = ttk.Combobox(frame_assuntos, state="readonly", width=18)
combo_disciplina_assunto.grid(row=1, column=1, sticky="w", padx=(4, 8))
combo_disciplina_assunto.bind("<<ComboboxSelected>>", _set_assuntos_disciplina)

ttk.Label(frame_assuntos, text="Assunto:").grid(row=1, column=2, sticky="w")
combo_assunto_assunto = ttk.Combobox(frame_assuntos, state="readonly", width=20)
combo_assunto_assunto.grid(row=1, column=3, sticky="w", padx=(4, 8))

ttk.Label(frame_assuntos, text="Mínimo:").grid(row=1, column=4, sticky="w")
entrada_minimo_assunto = ttk.Entry(frame_assuntos, width=6)
entrada_minimo_assunto.grid(row=1, column=5, sticky="w", padx=(4, 0))
entrada_minimo_assunto.insert(0, "0")

ttk.Button(frame_assuntos, text="2) Adicionar/Atualizar", command=adicionar_ou_atualizar_regra).grid(
    row=2, column=0, columnspan=2, sticky="w", pady=(6, 6)
)
ttk.Button(frame_assuntos, text="Remover selecionada", command=remover_regra_selecionada).grid(
    row=2, column=2, columnspan=2, sticky="w", pady=(6, 6)
)
ttk.Button(frame_assuntos, text="Limpar regras", command=limpar_regras).grid(
    row=2, column=4, columnspan=2, sticky="w", pady=(6, 6)
)

ignorar_assuntos_var = tk.BooleanVar(value=False)
ttk.Checkbutton(
    frame_assuntos,
    text="Relevar assuntos por disciplina",
    variable=ignorar_assuntos_var,
).grid(row=4, column=0, columnspan=3, sticky="w", pady=(6, 0))

tree_assuntos = ttk.Treeview(
    frame_assuntos,
    columns=("disciplina", "assunto", "minimo"),
    show="headings",
    height=8,
)
tree_assuntos.heading("disciplina", text="Disciplina")
tree_assuntos.heading("assunto", text="Assunto")
tree_assuntos.heading("minimo", text="Mínimo")
tree_assuntos.column("disciplina", width=140, anchor="w")
tree_assuntos.column("assunto", width=260, anchor="w")
tree_assuntos.column("minimo", width=70, anchor="center")
tree_assuntos.grid(row=5, column=0, columnspan=6, sticky="ew")
tree_assuntos.bind("<<TreeviewSelect>>", _selecionar_regra_tabela)
frame_assuntos.columnconfigure(3, weight=1)

ttk.Button(frame, text="Validar banco de questões", command=executar_validacao).grid(
    row=6, column=0, columnspan=3, pady=(0, 6)
)

frame_validacao = ttk.LabelFrame(frame, text="Resultado da validação", padding=8)
frame_validacao.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(0, 8))

lbl_resumo_validacao = ttk.Label(frame_validacao, text="Sem validação executada.")
lbl_resumo_validacao.grid(row=0, column=0, sticky="w", pady=(0, 6))

tree_validacao = ttk.Treeview(
    frame_validacao,
    columns=("tipo", "status", "detalhe"),
    show="headings",
    height=8,
)
tree_validacao.heading("tipo", text="Tipo")
tree_validacao.heading("status", text="Status")
tree_validacao.heading("detalhe", text="Detalhe")
tree_validacao.column("tipo", width=110, anchor="w")
tree_validacao.column("status", width=90, anchor="center")
tree_validacao.column("detalhe", width=520, anchor="w")
tree_validacao.grid(row=1, column=0, sticky="ew")
tree_validacao.tag_configure("aviso", background="#fff3b0")
tree_validacao.tag_configure("erro", background="#ffd6d6")
tree_validacao.tag_configure("ok", background="#dff6dd")
frame_validacao.columnconfigure(0, weight=1)

area_log = scrolledtext.ScrolledText(
    frame, height=12, width=60, state="disabled", font=("Courier", 9)
)
area_log.grid(row=8, column=0, columnspan=3, sticky="nsew")

btn_gerar = ttk.Button(frame, text="GERAR PROVA", command=executar_geracao)
btn_gerar.grid(row=9, column=0, columnspan=3, pady=10)

frame.columnconfigure(1, weight=1)
frame.columnconfigure(2, weight=1)

atualizar_painel_distribuicao()

app.mainloop()
