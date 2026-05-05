"""
archiver.py
-----------
Arquiva o lote atual de trabalho em uma pasta nomeada dentro de archives/.

Fluxo de uso:
    1. Rode main.py para gerar o relatório.
    2. Rode archiver.py — informe um nome para o lote.
    3. As pastas input_msgs/, temp_pdfs/ e output/ são movidas para
       archives/{nome}/ e esvaziadas, deixando o projeto limpo para o próximo lote.

Estrutura gerada:
    archives/
    └── {nome_do_lote}/
        ├── input_msgs/   ← e-mails .msg originais
        ├── temp_pdfs/    ← PDFs intermediários
        └── output/       ← relatórios finais (xlsx, csv, log)
"""

from __future__ import annotations

import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

# ─── Pastas que compõem um lote ─────────────────────────────────────────────
LOTE_DIRS = [
    Path("input_msgs"),
    Path("temp_pdfs"),
    Path("output"),
]

ARCHIVES_ROOT = Path("archives")


def _slugify(text: str) -> str:
    """Converte o nome informado em um nome de pasta seguro."""
    slug = re.sub(r"[^\w\-]", "_", text.strip())
    slug = re.sub(r"_+", "_", slug).strip("_")
    return slug


def _count_files(directory: Path) -> int:
    """Conta arquivos (não subpastas) dentro de um diretório."""
    if not directory.exists():
        return 0
    return sum(1 for f in directory.rglob("*") if f.is_file())


def _ask_name() -> str:
    """Solicita e valida o nome do lote interativamente."""
    print()
    print("=" * 60)
    print("  ARCHIVER — Salvar lote atual")
    print("=" * 60)

    # Mostra quantos arquivos existem em cada pasta
    total = 0
    for d in LOTE_DIRS:
        count = _count_files(d)
        total += count
        status = f"{count} arquivo(s)" if count else "vazia"
        print(f"  {d}/  →  {status}")

    if total == 0:
        print("\n  Nenhum arquivo encontrado nas pastas do lote.")
        print("  Execute main.py antes de arquivar.")
        sys.exit(0)

    print()

    # Sugestão de nome com data
    default = datetime.now().strftime("%Y-%m-%d")

    while True:
        raw = input(f"  Nome do lote [{default}]: ").strip()
        if not raw:
            raw = default

        slug = _slugify(raw)
        if not slug:
            print("  Nome inválido. Use letras, números, traços ou underscores.")
            continue

        dest = ARCHIVES_ROOT / slug
        if dest.exists():
            print(f"  Já existe um lote com o nome '{slug}' em archives/.")
            override = input("  Deseja sobrescrever? (s/N): ").strip().lower()
            if override != "s":
                continue
            shutil.rmtree(dest)

        return slug


def archive(lote_name: str | None = None) -> Path:
    """
    Move o conteúdo das pastas do lote para archives/{lote_name}/.

    Parameters
    ----------
    lote_name : str, opcional
        Nome do lote. Se None, solicita interativamente ao usuário.

    Returns
    -------
    Path
        Caminho da pasta de arquivo criada.
    """
    if lote_name is None:
        lote_name = _ask_name()

    dest_root = ARCHIVES_ROOT / lote_name
    dest_root.mkdir(parents=True, exist_ok=True)

    print()
    moved_total = 0

    for src_dir in LOTE_DIRS:
        if not src_dir.exists() or _count_files(src_dir) == 0:
            print(f"  [ignorado]  {src_dir}/  (vazia ou inexistente)")
            continue

        dest_dir = dest_root / src_dir.name
        dest_dir.mkdir(parents=True, exist_ok=True)

        moved = 0
        for item in sorted(src_dir.rglob("*")):
            if not item.is_file():
                continue
            relative = item.relative_to(src_dir)
            target = dest_dir / relative
            target.parent.mkdir(parents=True, exist_ok=True)
            shutil.move(str(item), str(target))
            moved += 1

        print(f"  [ok]  {src_dir}/  →  {dest_dir}/  ({moved} arquivo(s) movido(s))")
        moved_total += moved

        # Remove subpastas vazias que possam ter ficado
        for sub in sorted(src_dir.rglob("*"), reverse=True):
            if sub.is_dir():
                try:
                    sub.rmdir()
                except OSError:
                    pass

    # Recria as pastas base vazias para o próximo uso
    for src_dir in LOTE_DIRS:
        src_dir.mkdir(parents=True, exist_ok=True)

    print()
    print(f"  Lote '{lote_name}' arquivado em:  {dest_root.resolve()}")
    print(f"  Total: {moved_total} arquivo(s) movido(s).")
    print()
    print("  Pastas limpas e prontas para o próximo lote:")
    for d in LOTE_DIRS:
        print(f"    {d}/  (vazia)")
    print()

    return dest_root


if __name__ == "__main__":
    archive()
