"""
extractor_msg.py
----------------
Módulo responsável por varrer a pasta input_msgs/, abrir cada arquivo .msg
do Outlook e extrair apenas os anexos em formato PDF, salvando-os em temp_pdfs/.

Convenção de nomeação dos arquivos salvos:
    {slug_do_email}__{nome_original_do_anexo}.pdf

O slug é derivado do nome do arquivo .msg (sem extensão), com caracteres
especiais substituídos por underscore, evitando sobrescrita entre emails.
"""

from __future__ import annotations

import re
import logging
from pathlib import Path

import extract_msg
from tqdm import tqdm
from tqdm.contrib.logging import logging_redirect_tqdm

logger = logging.getLogger(__name__)

# ─── Caminhos padrão ────────────────────────────────────────────────────────
INPUT_DIR = Path("input_msgs")
OUTPUT_DIR = Path("temp_pdfs")


def _safe_slug(text: str) -> str:
    """Converte uma string arbitrária em um slug seguro para nomes de arquivo."""
    return re.sub(r"[^\w\-]", "_", text).strip("_")


def extract_pdf_attachments(
    input_dir: Path = INPUT_DIR,
    output_dir: Path = OUTPUT_DIR,
) -> list[Path]:
    """
    Percorre *input_dir* em busca de arquivos .msg e salva os anexos PDF em
    *output_dir*.

    Returns
    -------
    list[Path]
        Lista com os caminhos de todos os PDFs salvos nesta execução.
    """
    output_dir.mkdir(parents=True, exist_ok=True)

    msg_files = sorted(input_dir.glob("*.msg"))
    if not msg_files:
        logger.warning("Nenhum arquivo .msg encontrado em '%s'.", input_dir)
        return []

    saved_pdfs: list[Path] = []

    with logging_redirect_tqdm():
        for msg_path in tqdm(msg_files, desc="  📧  Processando .msg", unit="msg", dynamic_ncols=True, leave=True):
            email_slug = _safe_slug(msg_path.stem)

            try:
                with extract_msg.openMsg(str(msg_path)) as msg:
                    attachments = msg.attachments
                    if not attachments:
                        logger.info("  ↳ '%s' não possui anexos.", msg_path.name)
                        continue

                    for att in attachments:
                        original_name: str = att.longFilename or att.shortFilename or "anexo"
                        if not original_name.lower().endswith(".pdf"):
                            logger.debug(
                                "  ↳ Anexo ignorado (não-PDF): '%s'", original_name
                            )
                            continue

                        safe_name = _safe_slug(Path(original_name).stem)
                        dest_path = output_dir / f"{email_slug}__{safe_name}.pdf"

                        dest_path.write_bytes(att.data)
                        saved_pdfs.append(dest_path)
                        logger.info("  ↳ PDF salvo: '%s'", dest_path.name)

            except Exception as exc:  # noqa: BLE001
                logger.error("Erro ao processar '%s': %s", msg_path.name, exc)

    return saved_pdfs


if __name__ == "__main__":
    logging.basicConfig(level=logging.INFO, format="%(message)s")
    pdfs = extract_pdf_attachments()
    print(f"\n✅  {len(pdfs)} PDF(s) extraído(s) para '{OUTPUT_DIR}'.")
