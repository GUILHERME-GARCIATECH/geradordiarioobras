import json
import os
import sys
import threading
import traceback
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import ttkbootstrap as tb
from ttkbootstrap.constants import *
from ttkbootstrap.widgets import DateEntry

from main import gerar_relatorio, listar_obras, analisar_periodo_obra

try:
    import win32com.client  # type: ignore
    WORD_DISPONIVEL = True
except Exception:
    WORD_DISPONIVEL = False


APP_NOME = "GeradorDiarioObra"
VERSAO_APP = "1.3.1"
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
        "Obra_2 - Exemplo",
    ],
    "assumir_ultimo_duplicado": True,
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


def get_date_value(widget) -> str:
    try:
        return widget.entry.get().strip()
    except Exception:
        try:
            return widget.get().strip()
        except Exception:
            return ""


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
            except Exception:
                pass

    if isinstance(obras_fixas, list):
        obras_validas = [str(o).strip() for o in obras_fixas if str(o).strip()]
        if obras_validas:
            return obras_validas

    return []


def nome_arquivo_sugerido(extensao: str = ".docx") -> str:
    obra = combo_obra.get().strip() or "relatorio"
    medicao = entrada_medicao.get().strip() or "sem_medicao"
    nome_base = f"Diario_{obra}_{medicao}medicao".replace(" ", "_")
    return f"{nome_base}{extensao}"


def escolher_destino() -> str:
    tipos = [("Documento Word", "*.docx")]
    if WORD_DISPONIVEL:
        tipos.append(("PDF", "*.pdf"))

    return filedialog.asksaveasfilename(
        title="Salvar diário como",
        defaultextension=".docx",
        filetypes=tipos,
        initialfile=nome_arquivo_sugerido(".docx"),
    )


def atualizar_status_config() -> None:
    caminho_excel = str(config.get("excel_path", "")).strip()
    caminho_template = str(config.get("template_path", "")).strip()

    excel_ok = False
    if caminho_excel:
        excel_ok = resolver_caminho_programa(caminho_excel).exists()

    template_ok = False
    if caminho_template:
        template_ok = resolver_caminho_programa(caminho_template).exists()

    texto_base_var.set("Base conectada" if excel_ok else "Base não configurada")
    texto_template_var.set("Template configurado" if template_ok else "Template não configurado")
    texto_pdf_var.set(
        "Exportação em PDF disponível"
        if WORD_DISPONIVEL
        else "Exportação em PDF indisponível neste computador"
    )


def limpar_resumo() -> None:
    resumo_registros_var.set("—")
    resumo_dias_var.set("—")
    resumo_duplicidades_var.set("—")
    resumo_contratante_var.set("—")
    resumo_objeto_var.set("—")


def recarregar_obras(manter_atual: bool = True) -> None:
    global config
    selecionada = combo_obra.get().strip()

    config = carregar_config()
    obras = carregar_obras()
    combo_obra["values"] = obras

    if manter_atual and selecionada and selecionada in obras:
        combo_obra.set(selecionada)
    elif obras:
        combo_obra.current(0)
    else:
        combo_obra.set("")

    atualizar_status_config()
    atualizar_resumo()


def validar_campos() -> str | None:
    obra = combo_obra.get().strip()
    data_inicio = get_date_value(entrada_inicio)
    data_fim = get_date_value(entrada_fim)
    medicao = entrada_medicao.get().strip()
    data_assinatura = get_date_value(entrada_assinatura)

    if not obra:
        return "Selecione uma obra."
    if not data_inicio:
        return "Informe a data inicial."
    if not data_fim:
        return "Informe a data final."
    if not medicao:
        return "Informe a medição."
    if not data_assinatura:
        return "Informe a data da assinatura."
    return None


def atualizar_preview_nome(*_args) -> None:
    texto_nome_arquivo_var.set(nome_arquivo_sugerido(".docx"))


def atualizar_resumo(*_args) -> None:
    atualizar_preview_nome()

    obra = combo_obra.get().strip()
    data_inicio = get_date_value(entrada_inicio)
    data_fim = get_date_value(entrada_fim)

    resumo_obra_var.set(obra or "—")
    resumo_periodo_var.set(f"{data_inicio} a {data_fim}" if data_inicio and data_fim else "—")

    caminho_excel_txt = str(config.get("excel_path", "")).strip()
    if not caminho_excel_txt:
        resumo_status_var.set("Configure a base de dados nas configurações.")
        limpar_resumo()
        return

    if not obra or not data_inicio or not data_fim:
        resumo_status_var.set("Preencha os campos para visualizar o resumo.")
        limpar_resumo()
        return

    try:
        caminho_excel = str(resolver_caminho_programa(caminho_excel_txt))
        analise = analisar_periodo_obra(
            obra=obra,
            data_inicio_txt=data_inicio,
            data_fim_txt=data_fim,
            caminho_excel=caminho_excel,
        )

        resumo_registros_var.set(str(analise["total_registros"]))
        resumo_dias_var.set(str(analise["total_dias"]))
        resumo_duplicidades_var.set(str(analise["total_duplicidades"]))
        resumo_contratante_var.set(analise["contratante"] or "—")
        resumo_objeto_var.set(analise["objeto"] or "—")

        if analise["total_registros"] > 0:
            if analise["total_duplicidades"] > 0:
                resumo_status_var.set(
                    f"Foram encontrados registros. Duplicidades: {analise['total_duplicidades']} "
                    f"(o sistema usará o último registro do dia se essa opção estiver ativa)."
                )
            else:
                resumo_status_var.set("Resumo carregado com sucesso. Tudo pronto para gerar.")
        else:
            resumo_status_var.set("Nenhum registro encontrado para a obra e período selecionados.")

    except Exception as e:
        limpar_resumo()
        resumo_status_var.set(f"Não foi possível montar o resumo: {e}")


def set_loading(ativo: bool, texto: str = "") -> None:
    estado = "disabled" if ativo else "normal"

    for widget in widgets_bloqueaveis:
        try:
            widget.configure(state=estado)
        except Exception:
            pass

    if ativo:
        progresso.start(12)
        status_execucao_var.set(texto or "Gerando diário...")
        botao_gerar.configure(text="Gerando...", bootstyle="warning")
    else:
        progresso.stop()
        status_execucao_var.set("Pronto para gerar.")
        botao_gerar.configure(text="Gerar diário", bootstyle="success")


def abrir_configuracoes() -> None:
    global config

    janela_cfg = tb.Toplevel(janela)
    janela_cfg.title("Configurações")
    janela_cfg.geometry("860x560")
    janela_cfg.resizable(False, False)
    janela_cfg.transient(janela)
    janela_cfg.grab_set()

    frame_cfg = ttk.Frame(janela_cfg, padding=20)
    frame_cfg.pack(fill="both", expand=True)

    ttk.Label(
        frame_cfg,
        text="Configurações",
        font=("Segoe UI", 15, "bold"),
        bootstyle="inverse-secondary",
    ).grid(row=0, column=0, columnspan=4, sticky="w", pady=(0, 18))

    ttk.Label(frame_cfg, text="Planilha Excel", font=("Segoe UI", 10, "bold")).grid(
        row=1, column=0, sticky="w", pady=(0, 6)
    )
    var_excel = tk.StringVar(value=str(config.get("excel_path", "")).strip())
    entry_excel = ttk.Entry(frame_cfg, textvariable=var_excel, width=78)
    entry_excel.grid(row=2, column=0, columnspan=3, sticky="ew", padx=(0, 8), pady=(0, 12))

    def procurar_excel() -> None:
        caminho = filedialog.askopenfilename(
            title="Selecione a planilha Excel",
            filetypes=[("Arquivos Excel", "*.xlsx *.xlsm *.xls")],
        )
        if caminho:
            var_excel.set(caminho)

    ttk.Button(
        frame_cfg,
        text="Procurar",
        command=procurar_excel,
        bootstyle="secondary-outline",
    ).grid(row=2, column=3, sticky="e", pady=(0, 12))

    ttk.Label(frame_cfg, text="Template Word", font=("Segoe UI", 10, "bold")).grid(
        row=3, column=0, sticky="w", pady=(0, 6)
    )
    var_template = tk.StringVar(value=str(config.get("template_path", "")).strip())
    entry_template = ttk.Entry(frame_cfg, textvariable=var_template, width=78)
    entry_template.grid(row=4, column=0, columnspan=2, sticky="ew", padx=(0, 8), pady=(0, 12))

    def procurar_template() -> None:
        caminho = filedialog.askopenfilename(
            title="Selecione o template Word",
            filetypes=[("Documento Word", "*.docx")],
        )
        if caminho:
            var_template.set(caminho)

    def restaurar_template_padrao() -> None:
        var_template.set(str(TEMPLATE_PADRAO_REL))

    ttk.Button(
        frame_cfg,
        text="Procurar",
        command=procurar_template,
        bootstyle="secondary-outline",
    ).grid(row=4, column=2, sticky="ew", padx=(0, 8), pady=(0, 12))

    ttk.Button(
        frame_cfg,
        text="Usar padrão",
        command=restaurar_template_padrao,
        bootstyle="secondary-outline",
    ).grid(row=4, column=3, sticky="ew", pady=(0, 12))

    usar_excel_var = tk.BooleanVar(value=bool(config.get("usar_obras_do_excel", True)))
    ttk.Checkbutton(
        frame_cfg,
        text="Usar obras da planilha Excel",
        variable=usar_excel_var,
        bootstyle="round-toggle",
    ).grid(row=5, column=0, columnspan=4, sticky="w", pady=(4, 6))

    assumir_ultimo_var = tk.BooleanVar(value=bool(config.get("assumir_ultimo_duplicado", True)))
    ttk.Checkbutton(
        frame_cfg,
        text="Assumir último registro quando houver duplicidade",
        variable=assumir_ultimo_var,
        bootstyle="round-toggle",
    ).grid(row=6, column=0, columnspan=4, sticky="w", pady=(0, 14))

    ttk.Label(frame_cfg, text="Obras fixas (uma por linha)", font=("Segoe UI", 10, "bold")).grid(
        row=7, column=0, columnspan=4, sticky="w", pady=(0, 6)
    )

    txt_obras = tk.Text(
        frame_cfg,
        width=80,
        height=11,
        relief="solid",
        borderwidth=1,
        font=("Consolas", 10),
    )
    txt_obras.grid(row=8, column=0, columnspan=4, sticky="nsew")
    txt_obras.insert("1.0", "\n".join(config.get("obras_fixas", [])))

    rodape_cfg = ttk.Frame(frame_cfg)
    rodape_cfg.grid(row=9, column=0, columnspan=4, sticky="ew", pady=(18, 0))

    def salvar_configuracoes() -> None:
        global config

        excel_txt = var_excel.get().strip()
        template_txt = var_template.get().strip()
        usar_excel = usar_excel_var.get()
        assumir_ultimo = assumir_ultimo_var.get()
        obras_digitadas = txt_obras.get("1.0", "end").splitlines()
        obras_fixas = [obra.strip() for obra in obras_digitadas if obra.strip()]

        if usar_excel and not excel_txt:
            messagebox.showerror("Configuração inválida", "Selecione a planilha Excel.")
            return

        if excel_txt:
            excel_resolvido = resolver_caminho_programa(excel_txt)
            if not excel_resolvido.exists():
                messagebox.showerror("Planilha inválida", "O caminho da planilha informado não existe.")
                return

        if template_txt:
            template_resolvido = resolver_caminho_programa(template_txt)
            if not template_resolvido.exists():
                messagebox.showerror("Template inválido", "O caminho do template informado não existe.")
                return
            if template_resolvido.suffix.lower() != ".docx":
                messagebox.showerror("Template inválido", "O template precisa ser um arquivo .docx.")
                return

        if not usar_excel and not obras_fixas:
            messagebox.showerror(
                "Configuração inválida",
                "Se o Excel estiver desativado, informe pelo menos uma obra fixa.",
            )
            return

        config["excel_path"] = excel_txt
        config["template_path"] = template_txt or str(TEMPLATE_PADRAO_REL)
        config["usar_obras_do_excel"] = usar_excel
        config["assumir_ultimo_duplicado"] = assumir_ultimo
        config["obras_fixas"] = obras_fixas

        salvar_config(config)
        recarregar_obras()
        janela_cfg.destroy()
        messagebox.showinfo("Configurações", "Configurações salvas com sucesso.")

    ttk.Button(
        rodape_cfg,
        text="Cancelar",
        command=janela_cfg.destroy,
        bootstyle="secondary-outline",
    ).pack(side="right")

    ttk.Button(
        rodape_cfg,
        text="Salvar configurações",
        command=salvar_configuracoes,
        bootstyle="success",
    ).pack(side="right", padx=(0, 8))

    frame_cfg.columnconfigure(0, weight=1)
    frame_cfg.columnconfigure(1, weight=1)
    frame_cfg.columnconfigure(2, weight=0)
    frame_cfg.columnconfigure(3, weight=0)
    frame_cfg.rowconfigure(8, weight=1)


def finalizar_geracao_sucesso(caminho_final: Path) -> None:
    set_loading(False)
    atualizar_resumo()
    messagebox.showinfo(
        "Diário gerado",
        f"Diário gerado com sucesso.\n\nArquivo salvo em:\n{caminho_final}",
    )


def finalizar_geracao_erro(erro: Exception) -> None:
    set_loading(False)

    if isinstance(erro, FileNotFoundError):
        registrar_erro("Arquivo não encontrado", erro)
        messagebox.showerror(
            "Arquivo não encontrado",
            f"{erro}\n\nVerifique o caminho do Excel ou do template.",
        )
        return

    if isinstance(erro, ValueError):
        registrar_erro("Validação de dados", erro)
        messagebox.showerror("Erro de validação", str(erro))
        return

    registrar_erro("Erro inesperado ao gerar relatório", erro)
    messagebox.showerror(
        "Erro inesperado",
        f"Ocorreu um erro ao gerar o diário.\n\nDetalhes: {erro}\n\nLog salvo em:\n{ARQUIVO_LOG}",
    )


def executar() -> None:
    erro_validacao = validar_campos()
    if erro_validacao:
        messagebox.showwarning("Campos obrigatórios", erro_validacao)
        return

    destino = escolher_destino()
    if not destino:
        return

    obra = combo_obra.get().strip()
    data_inicio = get_date_value(entrada_inicio)
    data_fim = get_date_value(entrada_fim)
    medicao = entrada_medicao.get().strip()
    data_assinatura = get_date_value(entrada_assinatura)

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
            "Para exportar em PDF, o Microsoft Word precisa estar instalado neste computador.",
        )
        return

    set_loading(True, "Gerando diário...")

    def tarefa() -> None:
        try:
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

            janela.after(0, lambda: finalizar_geracao_sucesso(Path(caminho_final)))

        except Exception as e:
            janela.after(0, lambda: finalizar_geracao_erro(e))

    threading.Thread(target=tarefa, daemon=True).start()


janela = tb.Window(themename="flatly")
janela.title("Gerador de Diário de Obras")
janela.geometry("960x720")
janela.minsize(900, 680)

style = janela.style
style.configure("Titulo.TLabel", font=("Segoe UI", 20, "bold"))
style.configure("Subtitulo.TLabel", font=("Segoe UI", 10))
style.configure("InfoTitle.TLabel", font=("Segoe UI", 9))
style.configure("InfoValue.TLabel", font=("Segoe UI", 10, "bold"))
style.configure("Status.TLabel", font=("Segoe UI", 10))

container = ttk.Frame(janela, padding=22)
container.pack(fill=BOTH, expand=YES)

topo = ttk.Frame(container)
topo.pack(fill=X, pady=(0, 18))

ttk.Label(topo, text="Gerador de Diário de Obras", style="Titulo.TLabel").pack(side=LEFT)
ttk.Label(
    topo,
    text=f"v{VERSAO_APP}",
    bootstyle="secondary",
    padding=(10, 8),
).pack(side=LEFT, padx=(10, 0))

botao_config = ttk.Button(
    topo,
    text="Configurações",
    command=abrir_configuracoes,
    bootstyle="secondary-outline",
)
botao_config.pack(side=RIGHT)

ttk.Label(
    container,
    text="Preencha os dados principais, confira o resumo e gere o diário.",
    style="Subtitulo.TLabel",
    bootstyle="secondary",
).pack(anchor=W, pady=(0, 14))

card_dados = ttk.Labelframe(container, text="Dados do diário", padding=18, bootstyle="secondary")
card_dados.pack(fill=X, pady=(0, 14))

linha1 = ttk.Frame(card_dados)
linha1.pack(fill=X, pady=(0, 10))

ttk.Label(linha1, text="Obra").grid(row=0, column=0, sticky=W, padx=(0, 8))
combo_obra = ttk.Combobox(linha1, width=48, state="readonly", bootstyle="secondary")
combo_obra.grid(row=1, column=0, columnspan=3, sticky=EW, padx=(0, 16))

ttk.Label(linha1, text="Medição").grid(row=0, column=3, sticky=W, padx=(0, 8))
entrada_medicao = ttk.Entry(linha1, width=18)
entrada_medicao.grid(row=1, column=3, sticky=EW)

linha1.columnconfigure(0, weight=1)
linha1.columnconfigure(1, weight=1)
linha1.columnconfigure(2, weight=1)
linha1.columnconfigure(3, weight=0)

linha2 = ttk.Frame(card_dados)
linha2.pack(fill=X, pady=(0, 8))

ttk.Label(linha2, text="Data inicial").grid(row=0, column=0, sticky=W, padx=(0, 8))
entrada_inicio = DateEntry(
    linha2,
    width=18,
    dateformat="%d/%m/%Y",
    bootstyle="secondary",
)
entrada_inicio.grid(row=1, column=0, sticky=W, padx=(0, 16))

ttk.Label(linha2, text="Data final").grid(row=0, column=1, sticky=W, padx=(0, 8))
entrada_fim = DateEntry(
    linha2,
    width=18,
    dateformat="%d/%m/%Y",
    bootstyle="secondary",
)
entrada_fim.grid(row=1, column=1, sticky=W, padx=(0, 16))

ttk.Label(linha2, text="Data da assinatura").grid(row=0, column=2, sticky=W, padx=(0, 8))
entrada_assinatura = DateEntry(
    linha2,
    width=18,
    dateformat="%d/%m/%Y",
    bootstyle="secondary",
)
entrada_assinatura.grid(row=1, column=2, sticky=W)

card_status = ttk.Labelframe(container, text="Status da configuração", padding=18, bootstyle="info")
card_status.pack(fill=X, pady=(0, 14))

linha_status = ttk.Frame(card_status)
linha_status.pack(fill=X)

texto_base_var = tk.StringVar(value="—")
texto_template_var = tk.StringVar(value="—")
texto_pdf_var = tk.StringVar(value="—")
texto_nome_arquivo_var = tk.StringVar(value="—")

def bloco_info(parent, titulo: str, var: tk.StringVar) -> None:
    frame = ttk.Frame(parent)
    frame.pack(side=LEFT, fill=X, expand=YES, padx=(0, 16))
    ttk.Label(frame, text=titulo, style="InfoTitle.TLabel", bootstyle="secondary").pack(anchor=W)
    ttk.Label(frame, textvariable=var, style="InfoValue.TLabel").pack(anchor=W, pady=(2, 0))

bloco_info(linha_status, "Base de dados", texto_base_var)
bloco_info(linha_status, "Template", texto_template_var)
bloco_info(linha_status, "PDF", texto_pdf_var)
bloco_info(linha_status, "Nome sugerido", texto_nome_arquivo_var)

card_resumo = ttk.Labelframe(container, text="Resumo", padding=18, bootstyle="primary")
card_resumo.pack(fill=BOTH, expand=YES, pady=(0, 14))

grid_resumo = ttk.Frame(card_resumo)
grid_resumo.pack(fill=X)

resumo_obra_var = tk.StringVar(value="—")
resumo_periodo_var = tk.StringVar(value="—")
resumo_registros_var = tk.StringVar(value="—")
resumo_dias_var = tk.StringVar(value="—")
resumo_duplicidades_var = tk.StringVar(value="—")
resumo_contratante_var = tk.StringVar(value="—")
resumo_objeto_var = tk.StringVar(value="—")
resumo_status_var = tk.StringVar(value="Preencha os campos para visualizar o resumo.")

def campo_resumo(parent, titulo: str, var: tk.StringVar, row: int, col: int) -> None:
    bloco = ttk.Frame(parent)
    bloco.grid(row=row, column=col, sticky=EW, padx=10, pady=8)
    ttk.Label(bloco, text=titulo, style="InfoTitle.TLabel", bootstyle="secondary").pack(anchor=W)
    ttk.Label(bloco, textvariable=var, style="InfoValue.TLabel", wraplength=240).pack(anchor=W, pady=(2, 0))

campo_resumo(grid_resumo, "Obra", resumo_obra_var, 0, 0)
campo_resumo(grid_resumo, "Período", resumo_periodo_var, 0, 1)
campo_resumo(grid_resumo, "Registros encontrados", resumo_registros_var, 0, 2)
campo_resumo(grid_resumo, "Dias válidos", resumo_dias_var, 1, 0)
campo_resumo(grid_resumo, "Duplicidades", resumo_duplicidades_var, 1, 1)
campo_resumo(grid_resumo, "Contratante", resumo_contratante_var, 1, 2)
campo_resumo(grid_resumo, "Objeto", resumo_objeto_var, 2, 0)

grid_resumo.columnconfigure(0, weight=1)
grid_resumo.columnconfigure(1, weight=1)
grid_resumo.columnconfigure(2, weight=1)

ttk.Separator(card_resumo).pack(fill=X, pady=12)

ttk.Label(
    card_resumo,
    textvariable=resumo_status_var,
    style="Status.TLabel",
    wraplength=820,
    bootstyle="secondary",
).pack(anchor=W)

rodape = ttk.Frame(container)
rodape.pack(fill=X)

status_execucao_var = tk.StringVar(value="Pronto para gerar.")

progresso = ttk.Progressbar(rodape, mode="indeterminate", bootstyle="success-striped")
progresso.pack(fill=X, pady=(0, 10))

linha_acao = ttk.Frame(rodape)
linha_acao.pack(fill=X)

ttk.Label(
    linha_acao,
    textvariable=status_execucao_var,
    bootstyle="secondary",
).pack(side=LEFT)

botao_gerar = ttk.Button(
    linha_acao,
    text="Gerar diário",
    command=executar,
    bootstyle="success",
    width=22,
)
botao_gerar.pack(side=RIGHT)

widgets_bloqueaveis = [
    combo_obra,
    entrada_medicao,
    botao_gerar,
    botao_config,
]

entrada_medicao.bind("<KeyRelease>", atualizar_resumo)
combo_obra.bind("<<ComboboxSelected>>", atualizar_resumo)

try:
    entrada_inicio.entry.bind("<FocusOut>", atualizar_resumo)
    entrada_fim.entry.bind("<FocusOut>", atualizar_resumo)
    entrada_assinatura.entry.bind("<FocusOut>", atualizar_preview_nome)
except Exception:
    pass

atualizar_status_config()
recarregar_obras(manter_atual=False)
set_loading(False)
janela.mainloop()