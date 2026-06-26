import os
import pathlib
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
import pandas as pd
from dotenv import load_dotenv

# Carrega as credenciais seguras do arquivo .env
load_dotenv()
EMAIL_REMETENTE = os.getenv('EMAIL_REMETENTE')
SENHA_APP = os.getenv('SENHA_APP')

logging.basicConfig(level=logging.INFO, format='%(asctime)s - [%(levelname)s] - %(message)s')

CAMINHO_BASE = pathlib.Path(r'Bases de Dados')
CAMINHO_BACKUP = pathlib.Path(r'Backup Arquivos Lojas')

METAS = {
    'fat_dia': 2500.00,
    'fat_ano': 450000.00,
    'qtde_dia': 10,
    'qtde_ano': 30,
    'ticket_dia': 6.50,
    'ticket_ano': 6.50
}


def formatar_brl(valor):
    """Truque sênior: formata moeda no padrão brasileiro (1.234,56) sem usar a lib locale."""
    return f"R$ {valor:_.2f}".replace('.', ',').replace('_', '.')


def obter_farol(valor, meta):
    return '#2e7d32' if valor >= meta else '#c62828'


def extrair_e_tratar_dados():
    logging.info("Iniciando extração de dados...")
    try:
        df_emails = pd.read_excel(CAMINHO_BASE / 'Emails.xlsx')
        df_vendas = pd.read_excel(CAMINHO_BASE / 'Coffee_Shop_Sales.xlsx')

        df_vendas['Valor Final'] = df_vendas['transaction_qty'] * df_vendas['unit_price']

        df_vendas = df_vendas.rename(columns={
            'transaction_date': 'Data',
            'store_location': 'Loja',
            'product_category': 'Produto',
            'transaction_id': 'Código Venda'
        })

        df_vendas['Data'] = pd.to_datetime(df_vendas['Data'])
        return df_emails, df_vendas

    except Exception as erro:
        logging.critical(f"Falha crítica na leitura das bases: {erro}")
        return None, None


def gerar_snapshot_backup(df_loja, nome_loja, data_ref):
    try:
        pasta_dest = CAMINHO_BACKUP / nome_loja
        pasta_dest.mkdir(parents=True, exist_ok=True)

        caminho_completo = pasta_dest / f"{data_ref.strftime('%Y_%m_%d')}_{nome_loja}.xlsx"
        df_loja.to_excel(caminho_completo, index=False)
        return caminho_completo
    except Exception as erro:
        logging.error(f"Erro ao salvar backup de {nome_loja}: {erro}")
        return None


def enviar_email_gmail(destinatario, assunto, corpo_html, anexos=None):
    """Função core para envio de e-mails usando SMTP do Gmail."""
    if not EMAIL_REMETENTE or not SENHA_APP:
        logging.error("Credenciais não configuradas. Verifique o arquivo .env!")
        return False

    try:
        msg = MIMEMultipart()
        msg['From'] = EMAIL_REMETENTE
        msg['To'] = destinatario
        msg['Subject'] = assunto

        msg.attach(MIMEText(corpo_html, 'html'))

        # Adiciona anexos, se houver
        if anexos:
            for caminho in anexos:
                if caminho and os.path.exists(caminho):
                    with open(caminho, 'rb') as arquivo:
                        part = MIMEApplication(arquivo.read(), Name=os.path.basename(caminho))
                        part['Content-Disposition'] = f'attachment; filename="{os.path.basename(caminho)}"'
                        msg.attach(part)

        # Conecta ao servidor do Google
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(EMAIL_REMETENTE, SENHA_APP)
        server.send_message(msg)
        server.quit()

        return True
    except Exception as erro:
        logging.error(f"Erro na conexão SMTP: {erro}")
        return False


def despachar_relatorio_gerente(destinatario, nome_gerente, loja, data_ref, kpis, anexo_path):
    dia_str = data_ref.strftime('%d/%m/%Y')
    assunto = f"Audit Daily Store [{loja}] - {dia_str}"

    corpo_html = f"""
        <div style="font-family: Arial, sans-serif; color: #333;">
            <p>Olá, <strong>{nome_gerente}</strong>.</p>
            <p>Resultado operacional de ontem ({dia_str}) para a unidade <strong>{loja}</strong>:</p>

            <table border="1" cellpadding="6" cellspacing="0" style="border-collapse: collapse; width: 100%; max-width: 450px;">
                <tr style="background-color: #f8f9fa;">
                    <th style="text-align: left;">KPI</th>
                    <th>Realizado</th>
                    <th>Meta</th>
                    <th>Status</th>
                </tr>
                <tr>
                    <td>Faturamento</td>
                    <td align="center">{formatar_brl(kpis['fat_dia'])}</td>
                    <td align="center">{formatar_brl(METAS['fat_dia'])}</td>
                    <td align="center"><span style="color: {obter_farol(kpis['fat_dia'], METAS['fat_dia'])};">●</span></td>
                </tr>
                <tr>
                    <td>Mix Categorias</td>
                    <td align="center">{kpis['qtde_dia']}</td>
                    <td align="center">{METAS['qtde_dia']}</td>
                    <td align="center"><span style="color: {obter_farol(kpis['qtde_dia'], METAS['qtde_dia'])};">●</span></td>
                </tr>
                <tr>
                    <td>Ticket Médio</td>
                    <td align="center">{formatar_brl(kpis['ticket_dia'])}</td>
                    <td align="center">{formatar_brl(METAS['ticket_dia'])}</td>
                    <td align="center"><span style="color: {obter_farol(kpis['ticket_dia'], METAS['ticket_dia'])};">●</span></td>
                </tr>
            </table>
        </div>
    """

    anexos = [str(anexo_path)] if anexo_path else []
    return enviar_email_gmail(destinatario, assunto, corpo_html, anexos)


def gerar_rankings_gerais(df_vendas, data_ref):
    logging.info("Consolidando rankings gerais...")

    # 1. Ranking Anual
    df_rank_ano = df_vendas.groupby('Loja')['Valor Final'].sum().reset_index()
    df_rank_ano = df_rank_ano.sort_values(by='Valor Final', ascending=False)
    arq_ano = CAMINHO_BACKUP / f"{data_ref.strftime('%m_%d')}_Ranking_Anual.xlsx"
    df_rank_ano.to_excel(arq_ano, index=False)

    # 2. Ranking Dia
    df_vendas_dia = df_vendas[df_vendas['Data'] == data_ref]
    df_rank_dia = df_vendas_dia.groupby('Loja')['Valor Final'].sum().reset_index()
    df_rank_dia = df_rank_dia.sort_values(by='Valor Final', ascending=False)
    arq_dia = CAMINHO_BACKUP / f"{data_ref.strftime('%m_%d')}_Ranking_Dia.xlsx"
    df_rank_dia.to_excel(arq_dia, index=False)

    return df_rank_dia, df_rank_ano, arq_dia, arq_ano


def despachar_relatorio_diretoria(df_emails, df_rank_dia, df_rank_ano, data_ref, anexos):
    try:
        email_diretoria = df_emails.loc[df_emails['Loja'] == 'Diretoria', 'E-mail'].values[0]
    except IndexError:
        logging.error("E-mail da 'Diretoria' não encontrado na base de contatos!")
        return False

    dia_str = data_ref.strftime('%d/%m/%Y')
    assunto = f"Executive Summary [Coffee Shop Network] - {dia_str}"

    melhor_dia = df_rank_dia.iloc[0]
    pior_dia = df_rank_dia.iloc[-1]
    melhor_ano = df_rank_ano.iloc[0]
    pior_ano = df_rank_ano.iloc[-1]

    corpo_html = f"""
        <div style="font-family: Arial, sans-serif; color: #2b2b2b; max-width: 600px;">
            <h2 style="color: #1a237e; border-bottom: 2px solid #1a237e; padding-bottom: 5px;">Consolidado Operacional</h2>
            <p>Prezados membros da Diretoria, bom dia.</p>
            <p>Apresentamos os destaques de performance da rede no fechamento de <strong>{dia_str}</strong>:</p>

            <div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin: 15px 0;">
                <h3 style="margin-top:0; color: #333;">📍 Performance Diária ({dia_str})</h3>
                <ul style="margin-bottom:0;">
                    <li><strong>Melhor Unidade:</strong> {melhor_dia['Loja']} ({formatar_brl(melhor_dia['Valor Final'])})</li>
                    <li><strong>Menor Volume:</strong> {pior_dia['Loja']} ({formatar_brl(pior_dia['Valor Final'])})</li>
                </ul>
            </div>

            <div style="background-color: #f5f5f5; padding: 15px; border-radius: 5px; margin: 15px 0;">
                <h3 style="margin-top:0; color: #333;">🏆 Acumulado Anual (YTD)</h3>
                <ul style="margin-bottom:0;">
                    <li><strong>Líder da Rede:</strong> {melhor_ano['Loja']} ({formatar_brl(melhor_ano['Valor Final'])})</li>
                    <li><strong>Lanterna:</strong> {pior_ano['Loja']} ({formatar_brl(pior_ano['Valor Final'])})</li>
                </ul>
            </div>

            <p style="font-size: 13px; color: #555;">Os rankings completos detalhando todas as unidades seguem em anexo.</p>
            <p style="font-size: 12px; color: #888;">Rotina gerada automaticamente via Python Pipeline.</p>
        </div>
    """

    # Converte os caminhos pathlib para string antes de enviar
    anexos_str = [str(a) for a in anexos if a]
    return enviar_email_gmail(email_diretoria, assunto, corpo_html, anexos_str)


def main():
    logging.info("=== INICIANDO PIPELINE DE AUDITORIA E CONSOLIDAÇÃO ===")

    df_emails, df_vendas = extrair_e_tratar_dados()
    if df_vendas is None:
        return

    data_indicador = df_vendas['Data'].max()
    logging.info(f"Data de referência capturada: {data_indicador.strftime('%d/%m/%Y')}")

    # --- ETAPA 1: RELATÓRIOS DAS UNIDADES ---
    lojas_unicas = df_vendas['Loja'].unique()

    for loja in lojas_unicas:
        df_loja_total = df_vendas[df_vendas['Loja'] == loja]
        df_loja_dia = df_loja_total[df_loja_total['Data'] == data_indicador]

        caminho_anexo = gerar_snapshot_backup(df_loja_total, loja, data_indicador)
        kpis = {
            'fat_dia': df_loja_dia['Valor Final'].sum(),
            'qtde_dia': df_loja_dia['Produto'].nunique(),
            'ticket_dia': df_loja_dia.groupby('Código Venda')['Valor Final'].sum().mean()
        }

        try:
            email_dest = df_emails.loc[df_emails['Loja'] == loja, 'E-mail'].values[0]
            nome_ger = df_emails.loc[df_emails['Loja'] == loja, 'Gerente'].values[0]

            sucesso = despachar_relatorio_gerente(email_dest, nome_ger, loja, data_indicador, kpis, caminho_anexo)
            if sucesso:
                logging.info(f"[OK] Unidade {loja} notificada via Gmail.")
            else:
                logging.error(f"[FALHA] Não foi possível enviar o e-mail para a unidade {loja}.")
        except IndexError:
            logging.warning(f"Unidade '{loja}' sem e-mail cadastrado. Envio cancelado.")

    # --- ETAPA 2: FECHAMENTO DA DIRETORIA ---
    rank_dia, rank_ano, arq_dia, arq_ano = gerar_rankings_gerais(df_vendas, data_indicador)
    sucesso_diretoria = despachar_relatorio_diretoria(df_emails, rank_dia, rank_ano, data_indicador, [arq_dia, arq_ano])

    if sucesso_diretoria:
        logging.info(f"[OK] Relatório Executivo disparado com sucesso para a Diretoria.")
    else:
        logging.error(f"[FALHA] Não foi possível enviar o e-mail para a Diretoria.")

    logging.info("=== PIPELINE FINALIZADO ===")


if __name__ == '__main__':
    main()