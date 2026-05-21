from __future__ import annotations

from datetime import datetime, date


def parse_data(valor) -> date | None:
    if valor is None or valor == "":
        return None

    if isinstance(valor, datetime):
        return valor.date()

    if isinstance(valor, date):
        return valor

    texto = str(valor).strip()

    formatos = [
        "%d/%m/%Y",
        "%Y-%m-%d",
        "%d-%m-%Y",
    ]

    for fmt in formatos:
        try:
            return datetime.strptime(texto, fmt).date()
        except ValueError:
            continue

    return None


def filtrar_por_obra(registros: list[dict], obra_id_ou_texto: str) -> list[dict]:
    alvo = str(obra_id_ou_texto or "").strip().lower()
    if not alvo:
        return []

    resultado = []
    for r in registros:
        obra = str(r.get("obra") or "").strip().lower()
        obra_id = obra.split(" - ", 1)[0].strip()

        if alvo == obra or alvo == obra_id or obra.startswith(f"{alvo} - "):
            resultado.append(r)

    return resultado


def filtrar_por_periodo(
    registros: list[dict],
    data_inicio: date,
    data_fim: date,
    nome_coluna_data: str = "data",
) -> list[dict]:
    resultado = []

    for r in registros:
        data_registro = parse_data(r.get(nome_coluna_data))
        if not data_registro:
            continue

        if data_inicio <= data_registro <= data_fim:
            resultado.append(r)

    return resultado
