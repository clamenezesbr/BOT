# 🤖⚙️ BotPO

> Automação inteligente para leitura, extração e acumulação de dados de e-mails via Outlook/Exchange, gerando relatórios organizados automaticamente.

---

## 📌 Sobre o Projeto

O **BotPO** é um sistema automatizado em Python desenvolvido para **monitorar uma caixa de e-mails Outlook/Exchange**, extrair informações relevantes das mensagens recebidas e **acumular dados estruturados** (como pontos, registros ou indicadores) gerando relatórios em PDF de forma periódica.

Ideal para cenários onde é necessário consolidar informações que chegam por e-mail de forma recorrente — como boletins do SESI, notificações de sistemas, ou qualquer fonte de dados via e-mail.

---

## ✨ Funcionalidades

- 📥 **Leitura automática** de e-mails via Outlook/Exchange
- 🔍 **Extração de dados** estruturados do corpo das mensagens
- 📊 **Acumulação de pontos/registros** ao longo do tempo
- 📄 **Geração de relatórios em PDF** com os dados consolidados
- 🗂️ **Organização automática** de mensagens processadas
- 📝 **Log detalhado** de todas as execuções e erros encontrados

---

## 🗂️ Estrutura do Projeto

```
LeitorDeEmail/
│
├── 📁 logs/                         # Registros de execução do sistema
│   ├── acumulador_20260330_111751.log
│   ├── acumulador_20260330_120218.log
│   ├── acumulador_20260330_152333.log
│   ├── acumulador_20260330_152705.log
│   ├── acumulador_20260330_153012.log
│   ├── acumulador_20260330_154249.log  # Logs nomeados por data/hora de execução
│   └── log_erros.txt                # Arquivo centralizado de erros críticos
│
├── 📁 Mensagens/                    # E-mails capturados e organizados
│   └── (arquivos de mensagens processadas)
│
├── 📁 relatorios/                   # Relatórios gerados automaticamente em PDF
│   ├── relatorio_20260330_121054.pdf
│   └── relatorio_20260330_160835.pdf  # Nomeados com timestamp da geração
│
├── 📁 saida/                        # Dados de saída intermediários (ex: CSVs, JSONs)
│
└── 📁 scripts/                      # Código-fonte principal do projeto
    ├── acumulador_sesi.py           # Lógica de acumulação e processamento de dados
    ├── bot_extrator.py              # Bot de leitura e extração dos e-mails
    └── leitor.py                   # Módulo central de conexão e leitura do e-mail
```

### 📄 Descrição dos Scripts

| Arquivo | Função |
|---|---|
| `leitor.py` | Responsável pela **conexão com o Outlook/Exchange** e recuperação das mensagens da caixa de entrada |
| `bot_extrator.py` | Realiza o **parse e extração dos dados** relevantes do conteúdo de cada e-mail |
| `acumulador_sesi.py` | **Consolida e acumula** os dados extraídos, gerando os relatórios em PDF na pasta `relatorios/` |

---

## 🚀 Como Usar

### Pré-requisitos

- Python **3.8+**
- Conta **Outlook / Exchange** configurada localmente
- Bibliotecas necessárias (instalar via pip):

```bash
pip install -r requirements.txt
```

> ⚠️ Certifique-se de que o **Outlook está aberto e configurado** na máquina antes de executar.

### Execução

1. **Clone o repositório:**
```bash
git clone https://github.com/seu-usuario/LeitorDeEmail.git
cd LeitorDeEmail
```

2. **Execute o bot extrator** para capturar e processar os e-mails:
```bash
python scripts/bot_extrator.py
```

3. **Execute o acumulador** para consolidar os dados e gerar o relatório:
```bash
python scripts/acumulador_sesi.py
```

---

## 📋 Logs e Monitoramento

O sistema mantém registros detalhados de cada execução:

- Os arquivos em `logs/` são nomeados automaticamente com o padrão:
  ```
  acumulador_YYYYMMDD_HHMMSS.log
  ```
- O arquivo `log_erros.txt` centraliza todos os erros críticos encontrados durante as execuções, facilitando o diagnóstico de problemas.

---

## 📊 Relatórios

Os relatórios gerados ficam em `relatorios/` e seguem o padrão de nomenclatura:
```
relatorio_YYYYMMDD_HHMMSS.pdf
```
Cada relatório representa um snapshot dos dados acumulados até o momento da execução.

---

## 🛠️ Tecnologias Utilizadas

- **Python 3** — Linguagem principal
- **win32com / pywin32** — Integração com Outlook
- **ReportLab / FPDF** — Geração de PDFs
- **Logging** — Sistema de logs nativo do Python

---
