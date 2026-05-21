from __future__ import annotations

import ast
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]

SOURCE_FILES = [
    ROOT / "app.py",
    ROOT / "main.py",
    ROOT / "config.py",
    *sorted((ROOT / "src").glob("*.py")),
]

REQUIRED_FILES = [
    ROOT / "GeradorDiarioObra.spec",
    ROOT / "requirements.txt",
    ROOT / "requirements-build.txt",
    ROOT / "assets" / "icone.ico",
    ROOT / "assets" / "instalador_gerador_diario.iss",
    ROOT / "templates" / "modelopadrao.docx",
]


def check_required_files() -> None:
    missing = [path.relative_to(ROOT) for path in REQUIRED_FILES if not path.exists()]
    if missing:
        formatted = "\n".join(f"- {path}" for path in missing)
        raise SystemExit(f"Arquivos obrigatorios ausentes:\n{formatted}")


def check_python_syntax() -> None:
    for path in SOURCE_FILES:
        ast.parse(path.read_text(encoding="utf-8"), filename=str(path))


def main() -> None:
    check_required_files()
    check_python_syntax()
    print("Validacao do projeto concluida com sucesso.")


if __name__ == "__main__":
    main()
