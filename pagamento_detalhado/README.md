# Planilha de Pagamento Detalhado

Esta automação reúne informações de diversas fontes (Excel e PDFs) em uma planilha unificada, organizada e pronta para análise.  
O objetivo é economizar tempo no processo de consolidação de dados de pagamentos.

---

## ✨ Funcionalidades
- Copia dados de relatórios do sistema **Tela Web** (crédito e débito).
- Separa automaticamente as vendas por **bandeira** (Visa, Mastercard, Elo, Hipercard).
- Integra dados de **SGP** (crédito e débito).
- Extrai informações financeiras de **PDFs do Payware** (crédito e débito).
- Insere **imagens recortadas dos PDFs** diretamente nas planilhas.
- Renomeia arquivos finais de acordo com a data.
- Remove arquivos temporários após a execução.

---

## 📂 Estrutura esperada de arquivos
- `RELATÓRIO DE VENDAS ANALÍTICO - CRÉDITO.xlsx`
- `RELATÓRIO DE VENDAS ANALÍTICO - DÉBITO.xlsx`
- `SGP CRÉDITO.xlsx`
- `SGP DÉBITO.xlsx`
- `ANÁLISE BANDEIRAS - PAGAMENTO DETALHADO.xlsx`
- PDFs do Payware (ex.: `1 - CREDITO VISA - 4.14-Relatório Sintético de Movimentação de Estabelecimento.pdf`)

---

## 🛠️ Dependências
- `openpyxl`
- `pdfplumber`
- `PyMuPDF` (`fitz`)
- `Pillow`
- `re` (nativo do Python)
- `shutil` e `os` (nativos do Python)

---

## 🚀 Como executar
1. Coloque todos os arquivos necessários na mesma pasta do script.  
2. Execute o programa:
   ```bash
   python main.py
