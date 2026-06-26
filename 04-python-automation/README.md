Markdown

# 🐍 Automações e Scripts com Python

Coleção de scripts desenvolvidos para eliminar tarefas repetitivas, estruturar pipelines de dados e integrar serviços corporativos, focando em ganho de eficiência produtiva.

---

## 🚀 Projeto 1: [Daily Audit Pipeline]

## 🤖 Projeto em Destaque: Coffee Shop Network — Daily Audit Pipeline

> **O Problema:** Gestores de redes de varejo alimentício perdem, em média, 1h30 toda manhã extraindo relatórios do PDV, fatiando planilhas no Excel e redigindo e-mails individuais de cobrança de metas para cada gerente de filial.
>
> **A Solução:** Pipeline de automação ponta a ponta (`pipeline_auditoria.py`). O script lê a base centralizada de vendas transacionais, segmenta os dados por unidade, gera backups locais blindados contra falhas operacionais, calcula KPIs de faturamento e mix de produtos, aplica regra de farol (meta batida = verde / perdida = vermelho) e **dispara cards HTML formatados via SMTP (Gmail) para os gerentes, além de um Resumo Executivo consolidado para a Diretoria.**

### 📐 Arquitetura do Fluxo

1. **Ingestão:** Leitura dos datasets `Coffee_Shop_Sales.xlsx` e `Emails.xlsx`.
2. **Engenharia de Atributos:** Cálculo dinâmico de ticket médio e faturamento real transacional.
3. **Snapshot & Backup:** Geração automática de relatórios de auditoria `.xlsx` salvos em diretórios isolados por loja (`YYYY_MM_DD_Loja.xlsx`).
4. **Disparo Operacional:** Envio individualizado para o e-mail de cada Gerente contendo exclusivamente os indicadores da sua respectiva unidade.
5. **Disparo C-Level:** Geração de _Rankings de Performance (Dia vs. Ano)_ e despacho de Resumo Executivo estratégico para a mesa da Diretoria.

### 🛠️ Tecnologias e Bibliotecas Utilizadas

- **Pandas:** Manipulação, agregação e modelagem tabular de dados.
- **SMTP & Email.mime (Nativas):** Construção do corpo do e-mail em HTML, anexação de arquivos e comunicação multiplataforma com o servidor do Gmail.
- **Python-dotenv:** Gestão de credenciais de ambiente para garantir a segurança de senhas.
- **Pathlib:** Manipulação de caminhos de diretórios agnóstica de Sistema Operacional.
- **Logging:** Implementação de observabilidade e rastro de execução no terminal.

### ⚙️ Como Executar o Script

1. Clone este repositório em sua máquina local.
2. Instale as dependências necessárias rodando:
   ```bash
   pip install pandas openpyxl python-dotenv
   ```

---

## 🚀 Projeto 2: [Nome do Segundo Projeto]

## 🛠️ Projetos em Desenvolvimento

> 🚧 **Status:** Atualmente desenvolvendo o segundo script de automação e raspagem de dados em Python. Em breve, os códigos fontes e guias de execução estarão disponíveis aqui!
