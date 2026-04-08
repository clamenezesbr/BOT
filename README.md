# Automações Inteligentes

> Repositório central de bots desenvolvidos em Python para **automação de processos**, eliminação de tarefas repetitivas e otimização do fluxo de trabalho.

---

## 📌 Sobre este Repositório

Este repositório reúne uma coleção de bots independentes, cada um desenvolvido para resolver um problema específico. Todos são construídos com foco em **simplicidade, eficiência e facilidade de uso**.

---

## 🗂️ Bots Disponíveis

| Bot | Pasta | Descrição | Status |
|-----|-------|-----------|--------|
| 🤖⚙️ BotPO | [`/PO`](./PO) | Automação de leitura e extração de dados de e-mails Outlook | ✅ Ativo |
| 🎣🤖 Fishing Bot | [`/Fishing`](./FISHING) | Automação do minigame de pesca do FiveM via visão computacional | ✅ Ativo |

---

## 🤖⚙️ BotPO — Automação de E-mails Outlook

Bot desenvolvido para **monitorar automaticamente uma caixa de e-mails Outlook/Exchange**, extrair dados estruturados das mensagens e gerar relatórios consolidados em PDF.

**Principais funcionalidades:**
- 📥 Leitura automática de e-mails via Outlook
- 🔍 Extração e acumulação de dados das mensagens
- 📄 Geração de relatórios em PDF com timestamp
- 📝 Sistema de logs completo por execução

➡️ [Ver documentação completa do BotPO](./PO)

---

## 🎣🤖 Fishing Bot — Automação de Pesca FiveM

Bot desenvolvido para **automatizar o minigame de pesca em servidores FiveM**, utilizando visão computacional para rastrear o alvo em tempo real e relançar a pescaria de forma autônoma ao término de cada rodada.

**Principais funcionalidades:**
- 🖥️ Captura de tela de alta performance via `mss` (~120 fps)
- 🔍 Detecção do alvo por filtro de cor + análise de contornos com OpenCV
- 🖱️ Movimentação de cursor via `SendInput` (WinAPI) — compatível com jogo em foco
- 🔄 Relançamento automático da pescaria ao detectar fim do minigame
- 🐛 Modo debug com janela de visualização em tempo real
- 📦 Geração de executável `.exe` standalone com um único clique no `build.bat`

➡️ [Ver documentação completa do Fishing Bot](./FISHING)

---
