import json
import os
import sys
import traceback
import tkinter as tk
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from tkcalendar import DateEntry

from main import gerar_relatorio, listar_obras

try:
    import win32com.client  # type: ignore
    WORD_DISPONIVEL = True
except Exception:
    WORD_DISPONIVEL = False


APP_NOME = "GeradorDiarioObra"
VERSAO_APP = "1.2.0"
TEMPLATE_PADRAO_REL = Path("templates") / "modelopadrao.docx"


def pasta_base_programa() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resolver_caminho_programa(caminho: str | Path) -> Path:
    p = Path(caminho)
    if p.is_absolute():
        return p
    return pasta_base_programa() / p


def obter_pasta_app() -> Path:
    pasta = Path(os.getenv("LOCALAPPDATA", Path.home())) / APP_NOME
    pasta.mkdir(parents=True, exist_ok=True)
    return pasta


PASTA_APP = obter_pasta_app()
ARQUIVO_CONFIG = PASTA_APP / "config.json"
PASTA_LOGS = PASTA_APP / "logs"
PASTA_LOGS.mkdir(parents=True, exist_ok=True)
ARQUIVO_LOG = PASTA_LOGS / "erro.log"


CONFIG_PADRAO = {
    "excel_path": "",
    "template_path": str(TEMPLATE_PADRAO_REL),
    "usar_obras_do_excel": True,
    "obras_fixas": [
        "Obra_1 - Aldeia",
        "Obra_2 - Exemplo"
    ],
    "assumir_ultimo_duplicado": True
}


def corrigir_config_legado(config_lido: dict) -> dict:
    config_corrigido = CONFIG_PADRAO.copy()
    config_corrigido.update(config_lido or {})

    template_txt = str(config_corrigido.get("template_path", "")).strip()
    template_padrao_existe = resolver_caminho_programa(TEMPLATE_PADRAO_REL).exists()

    if not template_txt:
        config_corrigido["template_path"] = str(TEMPLATE_PADRAO_REL)
    else:
        caminho_resolvido = resolver_caminho_programa(template_txt)
        if not caminho_resolvido.exists() and template_padrao_existe:
            config_corrigido["template_path"] = str(TEMPLATE_PADRAO_REL)

    return config_corrigido


def carregar_config() -> dict:
    if not ARQUIVO_CONFIG.exists():
        config_novo = corrigir_config_legado({})
        salvar_config(config_novo)
        return config_novo.copy()

    try:
        with open(ARQUIVO_CONFIG, "r", encoding="utf-8") as f:
            config_lido = json.load(f)
    except Exception:
        config_lido = {}

    config_final = corrigir_config_legado(config_lido)
    salvar_config(config_final)
    return config_final


def salvar_config(config: dict) -> None:
    with open(ARQUIVO_CONFIG, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=4)


config = carregar_config()


def registrar_erro(contexto: str, erro: Exception) -> None:
    agora = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    with open(ARQUIVO_LOG, "a", encoding="utf-8") as f:
        f.write("=" * 80 + "\n")
        f.write(f"Data/Hora: {agora}\n")
        f.write(f"Contexto: {contexto}\n")
        f.write(f"Erro: {repr(erro)}\n")
        f.write(traceback.format_exc())
        f.write("\n")


def converter_docx_para_pdf(caminho_docx: str | Path, caminho_pdf: str | Path) -> None:
    if not WORD_DISPONIVEL:
        raise RuntimeError("Conversão para PDF requer Microsoft Word instalado.")

    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False

    try:
        doc = word.Documents.Open(str(Path(caminho_docx).resolve()))
        doc.SaveAs(str(Path(caminho_pdf).resolve()), FileFormat=17)
        doc.Close(False)
    finally:
        word.Quit()


def selecionar_excel():
    global config

    caminho = filedialog.askopenfilename(
        title="Selecione a planilha Excel",
        filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls")]
    )
    if caminho:
        config["excel_path"] = caminho
        salvar_config(config)
        atualizar_label_excel()
        recarregar_obras()
        messagebox.showinfo("Planilha definida", f"Planilha salva com sucesso:\n\n{caminho}")


def atualizar_label_excel():
    caminho = str(config.get("excel_path", "")).strip()
    if caminho:
        texto_excel_var.set(f"Excel: {caminho}")
    else:
        texto_excel_var.set("Excel: não configurado")


def atualizar_label_template():
    caminho = str(config.get("template_path", "")).strip()
    if caminho:
        texto_template_var.set(f"Template: {caminho}")
    else:
        texto_template_var.set("Template: não configurado")


def carregar_obras() -> list[str]:
    usar_excel = bool(config.get("usar_obras_do_excel", True))
    caminho_excel = str(config.get("excel_path", "")).strip()
    obras_fixas = config.get("obras_fixas", [])

    if usar_excel:
        caminho_excel_resolvido = None
        if caminho_excel:
            caminho_excel_resolvido = resolver_caminho_programa(caminho_excel)

        if caminho_excel_resolvido and caminho_excel_resolvido.exists():
            try:
                obras = listar_obras(caminho_excel_resolvido)
                if obras:
                    return obras
            except Exception as e:
                print(f"Erro ao carregar obras do Excel: {e}")

    if isinstance(obras_fixas, list):
        obras_validas = [str(o).strip() for o in obras_fixas if str(o).strip()]
        if obras_validas:
            return obras_validas

    return []


def recarregar_obras():
    global config
    config = carregar_config()
    obras = carregar_obras()
    combo_obra["values"] = obras
    if obras:
        combo_obra.current(0)
    else:
        combo_obra.set("")
    atualizar_label_excel()
    atualizar_label_template()


def escolher_destino() -> str:
    obra = combo_obra.get().strip() or "relatorio"
    medicao = entrada_medicao.get().strip() or "sem_medicao"
    nome_base = f"Diario_{obra}_{medicao}medicao".replace(" ", "_")

    tipos = [("Documento Word", "*.docx")]
    if WORD_DISPONIVEL:
        tipos.append(("PDF", "*.pdf"))

    caminho = filedialog.asksaveasfilename(
        title="Salvar relatório como",
        defaultextension=".docx",
        filetypes=tipos,
        initialfile=nome_base,
    )
    return caminho


def abrir_configuracoes():
    global config

    janela_cfg = tk.Toplevel(janela)
    janela_cfg.title("Configurações")
    janela_cfg.geometry("780x430")
    janela_cfg.resizable(False, False)
    janela_cfg.transient(janela)
    janela_cfg.grab_set()

    frame_cfg = ttk.Frame(janela_cfg, padding=20)
    frame_cfg.pack(fill="both", expand=True)

    ttk.Label(frame_cfg, text="⚙️ Configurações", font=("Segoe UI", 13, "bold")).grid(
        row=0, column=0, columnspan=4, sticky="w", pady=(0, 18)
    )

    ttk.Label(frame_cfg, text="Template Word:").grid(row=1, column=0, sticky="w", pady=8)
    var_template = tk.StringVar(value=str(config.get("template_path", "")).strip())
    entry_template = ttk.Entry(frame_cfg, textvariable=var_template, width=70)
    entry_template.grid(row=1, column=1, sticky="ew", pady=8, padx=(8, 8))

    def procurar_template():
        caminho = filedialog.askopenfilename(
            title="Selecione o template Word",
            filetypes=[("Documento Word", "*.docx")]
        )
        if caminho:
            var_template.set(caminho)

    def restaurar_template_padrao():
        var_template.set(str(TEMPLATE_PADRAO_REL))

    ttk.Button(frame_cfg, text="Procurar...", command=procurar_template).grid(
        row=1, column=2, sticky="e", pady=8
    )
    ttk.Button(frame_cfg, text="Restaurar padrão", command=restaurar_template_padrao).grid(
        row=1, column=3, sticky="e", pady=8, padx=(8, 0)
    )

    usar_excel_var = tk.BooleanVar(value=bool(config.get("usar_obras_do_excel", True)))
    ttk.Checkbutton(
        frame_cfg,
        text="Usar obras da planilha Excel",
        variable=usar_excel_var
    ).grid(row=2, column=0, columnspan=4, sticky="w", pady=(12, 6))

    assumir_ultimo_var = tk.BooleanVar(value=bool(config.get("assumir_ultimo_duplicado", True)))
    ttk.Checkbutton(
        frame_cfg,
        text="Assumir último registro quando houver duplicidade",
        variable=assumir_ultimo_var
    ).grid(row=3, column=0, columnspan=4, sticky="w", pady=6)

    ttk.Label(frame_cfg, text="Obras fixas (uma por linha):").grid(
        row=4, column=0, columnspan=4, sticky="w", pady=(16, 6)
    )

    txt_obras = tk.Text(frame_cfg, width=80, height=10)
    txt_obras.grid(row=5, column=0, columnspan=4, sticky="nsew")
    txt_obras.insert("1.0", "\n".join(config.get("obras_fixas", [])))

    botoes = ttk.Frame(frame_cfg)
    botoes.grid(row=6, column=0, columnspan=4, sticky="e", pady=(18, 0))

    def salvar_configuracoes():
        global config

        template_txt = var_template.get().strip()
        usar_excel = usar_excel_var.get()
        assumir_ultimo = assumir_ultimo_var.get()
        obras_digitadas = txt_obras.get("1.0", "end").splitlines()
        obras_fixas = [obra.strip() for obra in obras_digitadas if obra.strip()]

        if template_txt:
            template_resolvido = resolver_caminho_programa(template_txt)

            if not template_resolvido.exists():
                messagebox.showerror("Template inválido", "O caminho do template informado não existe.")
                return

            if template_resolvido.suffix.lower() != ".docx":
                messagebox.showerror("Template inválido", "O template precisa ser um arquivo .docx")
                return

        if not usar_excel and not obras_fixas:
            messagebox.showerror(
                "Configuração inválida",
                "Se o Excel estiver desativado, informe pelo menos uma obra fixa."
            )
            return

        config["template_path"] = template_txt or str(TEMPLATE_PADRAO_REL)
        config["usar_obras_do_excel"] = usar_excel
        config["assumir_ultimo_duplicado"] = assumir_ultimo
        config["obras_fixas"] = obras_fixas

        salvar_config(config)
        recarregar_obras()
        janela_cfg.destroy()
        messagebox.showinfo("Configurações", "Configurações salvas com sucesso.")

    ttk.Button(botoes, text="Salvar", command=salvar_configuracoes).pack(side="left", padx=(0, 8))
    ttk.Button(botoes, text="Cancelar", command=janela_cfg.destroy).pack(side="left")

    frame_cfg.columnconfigure(1, weight=1)
    frame_cfg.rowconfigure(5, weight=1)


def executar():
    try:
        obra = combo_obra.get().strip()
        data_inicio = entrada_inicio.get().strip()
        data_fim = entrada_fim.get().strip()
        medicao = entrada_medicao.get().strip()
        data_assinatura = entrada_assinatura.get().strip()

        if not obra or not data_inicio or not data_fim or not medicao or not data_assinatura:
            messagebox.showwarning("Campos obrigatórios", "Preencha todos os campos.")
            return

        destino = escolher_destino()
        if not destino:
            return

        botao_gerar.config(state="disabled")
        janela.update_idletasks()

        caminho_excel_txt = str(config.get("excel_path", "")).strip()
        caminho_excel = str(resolver_caminho_programa(caminho_excel_txt)) if caminho_excel_txt else None

        template_txt = str(config.get("template_path", "")).strip()
        template_path = str(resolver_caminho_programa(template_txt)) if template_txt else None

        assumir_ultimo = bool(config.get("assumir_ultimo_duplicado", True))

        destino_path = Path(destino)
        gerar_em_pdf = destino_path.suffix.lower() == ".pdf"

        if gerar_em_pdf and not WORD_DISPONIVEL:
            messagebox.showerror(
                "PDF indisponível",
                "Para exportar em PDF, o Microsoft Word precisa estar instalado neste computador."
            )
            return

        caminho_saida_docx = destino_path
        if gerar_em_pdf:
            caminho_saida_docx = destino_path.with_suffix(".docx")

        kwargs = {
            "obra": obra,
            "data_inicio_txt": data_inicio,
            "data_fim_txt": data_fim,
            "medicao": medicao,
            "data_assinatura": data_assinatura,
            "caminho_saida": caminho_saida_docx,
            "assumir_ultimo_duplicado": assumir_ultimo,
            "caminho_excel": caminho_excel,
            "caminho_template": template_path,
        }

        caminho_gerado = gerar_relatorio(**kwargs)

        caminho_final = caminho_gerado
        if gerar_em_pdf:
            converter_docx_para_pdf(caminho_gerado, destino_path)
            try:
                Path(caminho_gerado).unlink(missing_ok=True)
            except Exception:
                pass
            caminho_final = destino_path

        messagebox.showinfo(
            "Sucesso",
            f"Relatório gerado com sucesso!\n\n{caminho_final}"
        )

    except FileNotFoundError as e:
        registrar_erro("Arquivo não encontrado", e)
        messagebox.showerror(
            "Arquivo não encontrado",
            f"{e}\n\nVerifique o caminho do Excel ou do template."
        )

    except ValueError as e:
        registrar_erro("Validação de dados", e)
        messagebox.showerror("Erro de validação", str(e))

    except Exception as e:
        registrar_erro("Erro inesperado ao gerar relatório", e)
        messagebox.showerror(
            "Erro inesperado",
            f"Ocorreu um erro ao gerar o relatório.\n\nDetalhes: {e}\n\nLog salvo em:\n{ARQUIVO_LOG}"
        )

    finally:
        botao_gerar.config(state="normal")


janela = tk.Tk()
janela.title(f"Gerador de Diário de Obras - v{VERSAO_APP}")
janela.geometry("780x430")
janela.resizable(False, False)

frame = ttk.Frame(janela, padding=20)
frame.pack(fill="both", expand=True)

topo = ttk.Frame(frame)
topo.grid(row=0, column=0, columnspan=4, sticky="ew", pady=(0, 18))

titulo = ttk.Label(topo, text="Gerador de Diário de Obras", font=("Segoe UI", 14, "bold"))
titulo.pack(side="left")

ttk.Button(topo, text="⚙️", width=4, command=abrir_configuracoes).pack(side="right")

ttk.Label(frame, text="Obra:").grid(row=1, column=0, sticky="w", pady=6)
combo_obra = ttk.Combobox(frame, width=50, state="readonly")
combo_obra.grid(row=1, column=1, columnspan=3, sticky="ew", pady=6)

ttk.Label(frame, text="Data inicial:").grid(row=2, column=0, sticky="w", pady=6)
entrada_inicio = DateEntry(frame, width=18, date_pattern="dd/mm/yyyy", locale="pt_BR")
entrada_inicio.grid(row=2, column=1, sticky="w", pady=6)

ttk.Label(frame, text="Data final:").grid(row=3, column=0, sticky="w", pady=6)
entrada_fim = DateEntry(frame, width=18, date_pattern="dd/mm/yyyy", locale="pt_BR")
entrada_fim.grid(row=3, column=1, sticky="w", pady=6)

ttk.Label(frame, text="Número da medição:").grid(row=4, column=0, sticky="w", pady=6)
entrada_medicao = ttk.Entry(frame, width=20)
entrada_medicao.grid(row=4, column=1, sticky="w", pady=6)

ttk.Label(frame, text="Data da assinatura:").grid(row=5, column=0, sticky="w", pady=6)
entrada_assinatura = DateEntry(frame, width=18, date_pattern="dd/mm/yyyy", locale="pt_BR")
entrada_assinatura.grid(row=5, column=1, sticky="w", pady=6)

texto_excel_var = tk.StringVar()
label_excel = ttk.Label(frame, textvariable=texto_excel_var, wraplength=720)
label_excel.grid(row=6, column=0, columnspan=4, sticky="w", pady=(10, 4))

texto_template_var = tk.StringVar()
label_template = ttk.Label(frame, textvariable=texto_template_var, wraplength=720)
label_template.grid(row=7, column=0, columnspan=4, sticky="w", pady=(0, 8))

linha_botoes_cfg = ttk.Frame(frame)
linha_botoes_cfg.grid(row=8, column=0, columnspan=4, sticky="ew", pady=(0, 10))

ttk.Button(linha_botoes_cfg, text="Selecionar Excel", command=selecionar_excel).pack(side="left", padx=(0, 8))
ttk.Button(linha_botoes_cfg, text="Recarregar obras", command=recarregar_obras).pack(side="left")

botao_gerar = ttk.Button(frame, text="Gerar Relatório", command=executar)
botao_gerar.grid(row=9, column=0, columnspan=4, pady=(16, 0), ipadx=10, ipady=6)

if WORD_DISPONIVEL:
    texto_pdf = "PDF disponível"
else:
    texto_pdf = "PDF indisponível (Word não detectado neste PC)"

ttk.Label(frame, text=texto_pdf, wraplength=720).grid(
    row=10, column=0, columnspan=4, pady=(12, 0), sticky="w"
)

ttk.Label(frame, text=f"Config salvo em: {ARQUIVO_CONFIG}", wraplength=720).grid(
    row=11, column=0, columnspan=4, pady=(10, 0), sticky="w"
)

frame.columnconfigure(1, weight=1)
frame.columnconfigure(2, weight=1)
frame.columnconfigure(3, weight=1)

atualizar_label_excel()
atualizar_label_template()
recarregar_obras()

janela.mainloop()