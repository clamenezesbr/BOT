"""
============================
Extraí automaticamente todos os anexos .pdf de arquivos .msg (Outlook)
e os salva em uma pasta de saída, com log de erros e renomeação automática.
============================
"""

import os
import shutil
import extract_msg  # pip install extract-msg
from pathlib import Path

# ──────────────────────────────────────────────
# CONFIGURAÇÕES — ajuste aqui antes de rodar
# ──────────────────────────────────────────────
_BASE         = Path(__file__).parent.parent
PASTA_ENTRADA = str(_BASE / "LeitorDeEmail" / "Entrada")   # Pasta com os arquivos .msg
PASTA_SAIDA   = str(_BASE / "saida")                       # Pasta onde os PDFs serão salvos
ARQUIVO_LOG   = str(_BASE / "logs" / "log_erros.txt")      # Arquivo de log de erros/avisos
# ──────────────────────────────────────────────


def preparar_ambiente():
    """Garante que a pasta de saída existe e limpa/cria o arquivo de log."""
    if not os.path.exists(PASTA_SAIDA):
        os.makedirs(PASTA_SAIDA)
        print(f"[INFO] Pasta de saída criada: '{PASTA_SAIDA}'")
    else:
        print(f"[INFO] Pasta de saída já existe: '{PASTA_SAIDA}'")

    # Garante que a pasta de logs existe
    os.makedirs(os.path.dirname(ARQUIVO_LOG), exist_ok=True)

    # Inicia (ou sobrescreve) o log
    with open(ARQUIVO_LOG, "w", encoding="utf-8") as log:
        log.write("=== LOG DE EXECUÇÃO DO BOT EXTRATOR ===\n\n")


def gerar_nome_unico(pasta_destino: str, nome_original: str) -> str:
    """
    Se 'nome_original' já existir em 'pasta_destino', adiciona
    um sufixo numérico crescente: fatura.pdf → fatura_1.pdf → fatura_2.pdf …
    """
    destino_final = os.path.join(pasta_destino, nome_original)
    if not os.path.exists(destino_final):
        return destino_final

    nome_sem_ext, ext = os.path.splitext(nome_original)
    contador = 1
    while True:
        novo_nome = f"{nome_sem_ext}_{contador}{ext}"
        destino_final = os.path.join(pasta_destino, novo_nome)
        if not os.path.exists(destino_final):
            return destino_final
        contador += 1


def registrar_log(mensagem: str):
    """Acrescenta uma linha no arquivo de log."""
    with open(ARQUIVO_LOG, "a", encoding="utf-8") as log:
        log.write(mensagem + "\n")


def extrair_pdfs():
    """Loop principal: percorre todos os .msg e extrai anexos .pdf."""

    # 1. Listar arquivos .msg na pasta de entrada
    try:
        arquivos_msg = [
            f for f in os.listdir(PASTA_ENTRADA)
            if f.lower().endswith(".msg")
        ]
    except FileNotFoundError:
        print(f"[ERRO FATAL] Pasta de entrada não encontrada: '{PASTA_ENTRADA}'")
        registrar_log(f"[ERRO FATAL] Pasta '{PASTA_ENTRADA}' não encontrada.")
        return

    total = len(arquivos_msg)
    if total == 0:
        print("[AVISO] Nenhum arquivo .msg encontrado na pasta de entrada.")
        registrar_log("[AVISO] Nenhum arquivo .msg encontrado.")
        return

    print(f"\n[INFO] {total} arquivo(s) .msg encontrado(s). Iniciando extração...\n")
    print("─" * 55)

    # Contadores
    emails_ok        = 0
    emails_erro      = 0
    emails_sem_pdf   = 0
    total_pdfs_salvos = 0

    # 2. Iterar sobre cada arquivo
    for indice, nome_arquivo in enumerate(arquivos_msg, start=1):
        caminho_msg = os.path.join(PASTA_ENTRADA, nome_arquivo)
        print(f"[{indice:03d}/{total}] Processando: {nome_arquivo}")

        try:
            # 3a. Abrir o .msg em segundo plano (sem Outlook)
            with extract_msg.openMsg(caminho_msg) as msg:

                # 3b. Varrer os anexos do e-mail
                anexos = msg.attachments
                pdfs_neste_email = 0

                for anexo in anexos:
                    nome_anexo = anexo.longFilename or anexo.shortFilename or ""

                    # 3c. Filtro de extensão: apenas .pdf
                    if not nome_anexo.lower().endswith(".pdf"):
                        continue

                    # 3d. Gerar caminho de destino único (evita sobrescrita)
                    caminho_destino = gerar_nome_unico(PASTA_SAIDA, nome_anexo)

                    # Salvar o PDF
                    with open(caminho_destino, "wb") as arquivo_pdf:
                        arquivo_pdf.write(anexo.data)

                    print(f"         ✔ PDF salvo: {os.path.basename(caminho_destino)}")
                    pdfs_neste_email  += 1
                    total_pdfs_salvos += 1

                if pdfs_neste_email == 0:
                    print(f"         ℹ Nenhum PDF encontrado neste e-mail.")
                    registrar_log(f"[SEM PDF]  {nome_arquivo}")
                    emails_sem_pdf += 1
                else:
                    emails_ok += 1

        except Exception as erro:
            print(f"         ✘ ERRO ao processar: {erro}")
            registrar_log(f"[ERRO]     {nome_arquivo} → {erro}")
            emails_erro += 1

    # 4. Relatório final
    print("\n" + "═" * 55)
    print("  EXTRAÇÃO CONCLUÍDA!")
    print("═" * 55)
    print(f"  Arquivos .msg processados : {total}")
    print(f"  E-mails com PDF extraído  : {emails_ok}")
    print(f"  E-mails sem PDF           : {emails_sem_pdf}")
    print(f"  E-mails com erro          : {emails_erro}")
    print(f"  Total de PDFs salvos      : {total_pdfs_salvos}")
    print(f"  PDFs disponíveis em       : '{PASTA_SAIDA}/'")
    print(f"  Log de erros em           : '{ARQUIVO_LOG}'")
    print("═" * 55)

    registrar_log(
        f"\n--- RESUMO FINAL ---\n"
        f"Total .msg     : {total}\n"
        f"Com PDF        : {emails_ok}\n"
        f"Sem PDF        : {emails_sem_pdf}\n"
        f"Com erro       : {emails_erro}\n"
        f"PDFs salvos    : {total_pdfs_salvos}\n"
    )


# ──────────────────────────────────────────────
if __name__ == "__main__":
    preparar_ambiente()
    extrair_pdfs()
