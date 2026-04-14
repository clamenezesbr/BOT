"""
extractor.py
------------
Módulo responsável por iterar sobre os arquivos .msg em input_msgs/,
extrair apenas os anexos .pdf e salvá-los em temp_pdfs/.
"""

import logging
from pathlib import Path
from typing import List, Tuple

import extract_msg

logger = logging.getLogger(__name__)


def _ensure_dirs(input_dir: Path, output_dir: Path) -> None:
    """Garante que os diretórios necessários existam."""
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)


def extract_pdfs_from_msgs(
    input_dir: Path,
    output_dir: Path,
) -> Tuple[List[Path], int, int]:
    """
    Percorre todos os arquivos .msg em *input_dir*, extrai os anexos .pdf
    e os salva em *output_dir*.

    Parameters
    ----------
    input_dir : Path
        Pasta contendo os arquivos .msg.
    output_dir : Path
        Pasta de destino para os PDFs extraídos.

    Returns
    -------
    extracted_paths : list[Path]
        Lista com os caminhos dos PDFs salvos.
    total_msgs : int
        Total de arquivos .msg encontrados.
    skipped : int
        Total de .msg sem nenhum anexo PDF válido.
    """
    _ensure_dirs(input_dir, output_dir)

    msg_files: List[Path] = sorted(input_dir.glob("*.msg"))
    extracted_paths: List[Path] = []
    skipped: int = 0
    total_msgs: int = len(msg_files)

    if total_msgs == 0:
        logger.warning("Nenhum arquivo .msg encontrado em '%s'.", input_dir)
        return extracted_paths, total_msgs, skipped

    for msg_path in msg_files:
        logger.info("Processando .msg: %s", msg_path.name)
        try:
            msg = extract_msg.Message(str(msg_path))
        except Exception as exc:
            logger.error("Falha ao abrir '%s': %s", msg_path.name, exc)
            skipped += 1
            continue

        try:
            pdf_attachments = [
                att
                for att in msg.attachments
                if att.longFilename and att.longFilename.lower().endswith(".pdf")
            ]
        except Exception as exc:
            logger.error(
                "Erro ao listar anexos de '%s': %s", msg_path.name, exc
            )
            skipped += 1
            continue

        if not pdf_attachments:
            logger.warning(
                "'%s' não possui anexos .pdf — ignorado.", msg_path.name
            )
            skipped += 1
            continue

        for att in pdf_attachments:
            dest_path = output_dir / att.longFilename
            # Evita sobrescrita silenciosa: adiciona sufixo se já existir
            if dest_path.exists():
                stem = dest_path.stem
                suffix = dest_path.suffix
                counter = 1
                while dest_path.exists():
                    dest_path = output_dir / f"{stem}_{counter}{suffix}"
                    counter += 1

            try:
                dest_path.write_bytes(att.data)
                logger.info("PDF extraído: %s → %s", att.longFilename, dest_path)
                extracted_paths.append(dest_path)
            except Exception as exc:
                logger.error(
                    "Falha ao salvar PDF '%s': %s", att.longFilename, exc
                )

    return extracted_paths, total_msgs, skipped
