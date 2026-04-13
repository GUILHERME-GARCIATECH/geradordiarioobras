from __future__ import annotations

import sys
from collections import defaultdict
from datetime import datetime
from pathlib import Path

from config import (
    ARQUIVO_EXCEL,
    ARQUIVO_TEMPLATE,
    MAPA_ETAPAS,
    NOME_ABA_CADASTRO,
    NOME_ABA_RESPOSTAS,
)
from src.excel_reader import carregar_workbook, ler_aba_como_dicts
from src.filtros import filtrar_por_obra, filtrar_por_periodo, parse_data
from src.tarefas import extrair_tarefas
from src.diario_builder import buscar_cadastro_obra, montar_diario
from src.word_generator import gerar_docx_template


def pasta_base_programa() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def resolver_caminho_base(caminho: str | Path) -> Path:
    p = Path(caminho)
    if p.is_absolute():
        return p
    return pasta_base_programa() / p


def resolver_caminho_excel(caminho_excel: str | Path | None = None) -> Path:
    caminho_final = resolver_caminho_base(caminho_excel) if caminho_excel else resolver_caminho_base(ARQUIVO_EXCEL)

    if not caminho_final.exists():
        raise FileNotFoundError(f"Arquivo Excel não encontrado:\n{caminho_final}")

    return caminho_final


def resolver_caminho_template(caminho_template: str | Path | None = None) -> Path:
    caminho_final = resolver_caminho_base(caminho_template) if caminho_template else resolver_caminho_base(ARQUIVO_TEMPLATE)

    if not caminho_final.exists():
        raise FileNotFoundError(f"Template Word não encontrado:\n{caminho_final}")

    return caminho_final


def pedir_entrada_usuario():
    obra = input("Digite a obra (ex: Obra_3): ").strip()
    data_inicio = input("Data inicial (dd/mm/aaaa): ").strip()
    data_fim = input("Data final (dd/mm/aaaa): ").strip()
    medicao = input("Número da medição: ").strip()
    data_assinatura = input("Data da assinatura (dd/mm/aaaa): ").strip()
    return obra, data_inicio, data_fim, medicao, data_assinatura


def ordenar_registros(registros: list[dict]) -> list[dict]:
    def chave_ordenacao(registro: dict):
        data_reg = parse_data(registro.get("data"))
        hora_fim = str(registro.get("Hora de conclusão") or "").strip()
        hora_ini = str(registro.get("Hora de início") or "").strip()
        return (data_reg or datetime.min.date(), hora_fim, hora_ini)

    return sorted(registros, key=chave_ordenacao)


def agrupar_registros_por_data(registros: list[dict], nome_coluna_data: str = "data") -> dict:
    grupos = defaultdict(list)
    registros_ordenados = ordenar_registros(registros)

    for registro in registros_ordenados:
        dt = parse_data(registro.get(nome_coluna_data))
        if dt:
            grupos[dt].append(registro)

    return dict(sorted(grupos.items(), key=lambda x: x[0]))


def detectar_duplicatas_por_chave(registros: list[dict]) -> dict:
    grupos = defaultdict(list)

    for registro in registros:
        obra = str(registro.get("obra") or "").strip()
        data = parse_data(registro.get("data"))

        if obra and data:
            chave = (obra, data)
            grupos[chave].append(registro)

    return {chave: itens for chave, itens in grupos.items() if len(itens) > 1}


def resolver_duplicatas_amigavel(registros_por_data: dict, assumir_ultimo: bool = True) -> dict:
    resultado = {}

    for data_ref, registros_do_dia in registros_por_data.items():
        if len(registros_do_dia) == 1:
            resultado[data_ref] = registros_do_dia
            continue

        if assumir_ultimo:
            resultado[data_ref] = [registros_do_dia[-1]]
            continue

        print("\n========================================")
        print(f"AVISO: existem {len(registros_do_dia)} registros para o dia {data_ref}.")
        print("Isso indica duplicidade para a mesma obra/data.")
        escolha = input("Deseja continuar usando o ULTIMO registro desse dia? (s/n): ").strip().lower()

        if escolha != "s":
            print("Geração cancelada para revisão da planilha.")
            return {}

        resultado[data_ref] = [registros_do_dia[-1]]

    return resultado


def listar_obras(caminho_excel: str | Path | None = None) -> list[str]:
    caminho_excel_final = resolver_caminho_excel(caminho_excel)

    workbook = carregar_workbook(caminho_excel_final)
    respostas = ler_aba_como_dicts(workbook, NOME_ABA_RESPOSTAS)

    obras = []
    for r in respostas:
        obra = str(r.get("obra") or "").strip()
        if obra and obra not in obras:
            obras.append(obra)

    return sorted(obras)


def gerar_relatorio(
    obra: str,
    data_inicio_txt: str,
    data_fim_txt: str,
    medicao: str,
    data_assinatura: str,
    caminho_saida: str | Path,
    assumir_ultimo_duplicado: bool = True,
    caminho_excel: str | Path | None = None,
    caminho_template: str | Path | None = None,
) -> Path:
    data_inicio = datetime.strptime(data_inicio_txt, "%d/%m/%Y").date()
    data_fim = datetime.strptime(data_fim_txt, "%d/%m/%Y").date()
    periodo_texto = f"{data_inicio_txt} a {data_fim_txt}"

    caminho_excel_final = resolver_caminho_excel(caminho_excel)
    caminho_template_final = resolver_caminho_template(caminho_template)

    workbook = carregar_workbook(caminho_excel_final)

    respostas = ler_aba_como_dicts(workbook, NOME_ABA_RESPOSTAS)
    cadastros = ler_aba_como_dicts(workbook, NOME_ABA_CADASTRO)

    registros_obra = filtrar_por_obra(respostas, obra)
    registros_periodo = filtrar_por_periodo(registros_obra, data_inicio, data_fim)

    if not registros_periodo:
        raise ValueError("Nenhum registro encontrado para essa obra e período.")

    duplicatas = detectar_duplicatas_por_chave(registros_periodo)
    if duplicatas and not assumir_ultimo_duplicado:
        raise ValueError("Foram encontradas duplicatas para a mesma obra/data.")

    cadastro_obra = buscar_cadastro_obra(cadastros, registros_periodo[0].get("obra", ""))

    registros_por_data = agrupar_registros_por_data(registros_periodo)
    registros_por_data = resolver_duplicatas_amigavel(
        registros_por_data,
        assumir_ultimo=assumir_ultimo_duplicado,
    )

    if not registros_por_data:
        raise ValueError("Geração cancelada por duplicidade.")

    diarios = []

    for data_ref, registros_do_dia in registros_por_data.items():
        registro_base = registros_do_dia[-1]
        tarefas_consolidadas = extrair_tarefas(registro_base, MAPA_ETAPAS, limite=9)

        diario = montar_diario(
            registros_filtrados=registros_do_dia,
            cadastro_obra=cadastro_obra,
            periodo_texto=periodo_texto,
            medicao=medicao,
            data_assinatura=data_assinatura,
            tarefas_por_registro=tarefas_consolidadas,
            data_referencia=data_ref,
        )
        diarios.append(diario)

    if not diarios:
        raise ValueError("Nenhum diário pôde ser montado.")

    caminho_saida = Path(caminho_saida)
    caminho_saida.parent.mkdir(parents=True, exist_ok=True)

    gerar_docx_template(caminho_template_final, caminho_saida, diarios)

    return caminho_saida


def main(
    obra,
    data_inicio_txt,
    data_fim_txt,
    medicao,
    data_assinatura,
    caminho_excel: str | Path | None = None,
    caminho_template: str | Path | None = None,
):
    nome_arquivo = f"Diario_{obra}_{medicao}medicao.docx".replace(" ", "_")
    caminho_saida = Path("saida") / nome_arquivo
    return gerar_relatorio(
        obra=obra,
        data_inicio_txt=data_inicio_txt,
        data_fim_txt=data_fim_txt,
        medicao=medicao,
        data_assinatura=data_assinatura,
        caminho_saida=caminho_saida,
        assumir_ultimo_duplicado=False,
        caminho_excel=caminho_excel,
        caminho_template=caminho_template,
    )


if __name__ == "__main__":
    obra, data_inicio, data_fim, medicao, data_assinatura = pedir_entrada_usuario()
    caminho = main(obra, data_inicio, data_fim, medicao, data_assinatura)
    print(f"Documento gerado com sucesso em:\n{caminho}")