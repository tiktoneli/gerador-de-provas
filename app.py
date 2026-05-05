import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from gerador import (
    carregar_banco,
    carregar_config,
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


# ── helpers ───────────────────────────────────────────────────────────────────

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
    """Reads percentage fields from the distribution panel.
    Returns None if input is invalid (caller shows error)."""
    modelo_nome = combo_banca.get()
    dif_range = config[modelo_nome].get("dificuldade_range", [1, 5])
    resultado: dict[int, float] = {}
    for nivel in range(dif_range[0], dif_range[1] + 1):
        txt = dist_vars[nivel].get().strip().replace(",", ".")
        try:
            resultado[nivel] = float(txt) / 100.0
        except ValueError:
            return None
    return resultado


# ── painel de distribuição ────────────────────────────────────────────────────

def atualizar_painel_distribuicao(event=None) -> None:
    modelo_nome = combo_banca.get()
    modelo = config[modelo_nome]
    dif_range = modelo.get("dificuldade_range", [1, 5])
    dist = distribuicao_efetiva(modelo)

    lbl_range.configure(text=f"Faixa válida: {dif_range[0]} – {dif_range[1]}")

    for nivel in range(1, 6):
        in_range = dif_range[0] <= nivel <= dif_range[1]
        if in_range:
            dist_vars[nivel].set(f"{dist.get(nivel, 0) * 100:.1f}")
            dist_labels[nivel].grid()
            dist_entries[nivel].grid()
            dist_pct_labels[nivel].grid()
        else:
            dist_labels[nivel].grid_remove()
            dist_entries[nivel].grid_remove()
            dist_pct_labels[nivel].grid_remove()

    btn_reset_dist.configure(command=lambda: atualizar_painel_distribuicao())


# ── ações ─────────────────────────────────────────────────────────────────────

def executar_validacao() -> None:
    csv_path = entrada_csv.get().strip()
    if not csv_path:
        messagebox.showwarning("Aviso", "Selecione o banco de questões (CSV) primeiro.")
        return
    dist = _ler_distribuicao_gui()
    if dist is None:
        messagebox.showwarning("Aviso", "Verifique os valores da distribuição de dificuldades.")
        return
    try:
        df = carregar_banco(csv_path)
        modelo_nome = combo_banca.get()
        resultado = validar_banco(df, modelo_nome, config, dist)
        erros = resultado["erros"]
        avisos = resultado["avisos"]
        if erros:
            log("ERRO — questoes insuficientes (bloqueia geracao):")
            for e in erros:
                log(f"  {e}")
        if avisos:
            log("Aviso — distribuicao aproximada (geracao prossegue):")
            for a in avisos:
                log(f"  {a}")
        if not erros and not avisos:
            total = df.shape[0]
            log(f"Banco valido para {modelo_nome.upper()} — {total} questoes carregadas.")
    except Exception as e:
        log(f"✗ Erro na validação: {e}")


def _thread_gerar(
    csv_path: str,
    modelo_nome: str,
    semente: int | None,
    dist: dict[int, float],
) -> None:
    log(
        f"\n→ Gerando {modelo_nome.upper()} | "
        f"semente: {semente if semente is not None else 'aleatória'}"
    )
    try:
        arquivos = gerar_prova(modelo_nome, csv_path, semente, dist)
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

    dist = _ler_distribuicao_gui()
    if dist is None:
        messagebox.showwarning("Aviso", "Verifique os valores da distribuição de dificuldades.")
        return

    total_pct = sum(dist.values()) * 100
    if abs(total_pct - 100.0) > 0.5:
        messagebox.showwarning(
            "Aviso",
            f"A distribuição soma {total_pct:.1f}% — deve ser 100%.",
        )
        return

    semente_txt = entrada_semente.get().strip()
    semente = int(semente_txt) if semente_txt.isdigit() else None
    modelo_nome = combo_banca.get()

    btn_gerar.configure(state="disabled")
    threading.Thread(
        target=_thread_gerar,
        args=(csv_path, modelo_nome, semente, dist),
        daemon=True,
    ).start()


# ── GUI ───────────────────────────────────────────────────────────────────────

app = tk.Tk()
app.title("Gerador de Provas")
app.resizable(False, False)

frame = ttk.Frame(app, padding=16)
frame.pack(fill="both", expand=True)

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
frame_dist = ttk.LabelFrame(frame, text="Distribuição de dificuldades", padding=8)
frame_dist.grid(row=3, column=0, columnspan=3, sticky="ew", pady=6)

lbl_range = ttk.Label(frame_dist, text="")
lbl_range.grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 4))

dist_vars: dict[int, tk.StringVar] = {}
dist_labels: dict[int, ttk.Label] = {}
dist_entries: dict[int, ttk.Entry] = {}
dist_pct_labels: dict[int, ttk.Label] = {}

for i, nivel in enumerate(range(1, 6)):
    dist_vars[nivel] = tk.StringVar()

    lbl = ttk.Label(frame_dist, text=f"Nível {nivel}  {NOMES_NIVEL[nivel]}:")
    lbl.grid(row=i + 1, column=0, sticky="w", pady=1, padx=(0, 6))
    dist_labels[nivel] = lbl

    ent = ttk.Entry(frame_dist, textvariable=dist_vars[nivel], width=8)
    ent.grid(row=i + 1, column=1, sticky="w")
    dist_entries[nivel] = ent

    pct_lbl = ttk.Label(frame_dist, text="%")
    pct_lbl.grid(row=i + 1, column=2, sticky="w", padx=(2, 0))
    dist_pct_labels[nivel] = pct_lbl

btn_reset_dist = ttk.Button(frame_dist, text="Restaurar padrão", width=16)
btn_reset_dist.grid(row=7, column=0, columnspan=3, pady=(8, 0), sticky="w")

ttk.Separator(frame, orient="horizontal").grid(
    row=4, column=0, columnspan=3, sticky="ew", pady=10
)

ttk.Button(frame, text="Validar banco de questões", command=executar_validacao).grid(
    row=5, column=0, columnspan=3, pady=(0, 6)
)

area_log = scrolledtext.ScrolledText(
    frame, height=12, width=60, state="disabled", font=("Courier", 9)
)
area_log.grid(row=6, column=0, columnspan=3)

btn_gerar = ttk.Button(frame, text="GERAR PROVA", command=executar_geracao)
btn_gerar.grid(row=7, column=0, columnspan=3, pady=10)

# Inicializa painel com a primeira banca
atualizar_painel_distribuicao()

app.mainloop()
