import win32com.client
import pandas as pd
import shutil
from pathlib import Path
from io import StringIO
import logging
import re
from datetime import datetime
import sys
import warnings
import os
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl import Workbook

# Suprimir avisos
warnings.filterwarnings('ignore')

# --- 1. CONFIGURAÇÕES GERAIS ---
CAMINHO_PASTA_LOCAL = Path(r"C:\Users\matheus.augusto\OneDrive - Grupo Fleury\Planejamento Financeiro - TI&Telecom _ - Documentos\RELATÓRIO\Automações\Gerar Relatório OPEX")

# Data de Corte (Ignora e-mails antigos)
DATA_INICIO_LEITURA = datetime(2026, 1, 1)

if not Path.exists(CAMINHO_PASTA_LOCAL):
    try:
        CAMINHO_PASTA_LOCAL.mkdir(parents=True, exist_ok=True)
        logging.info(f"Pasta criada em: {CAMINHO_PASTA_LOCAL}")
        print(f"Pasta criada em: {CAMINHO_PASTA_LOCAL}")
    except:
        pass

ARQUIVO_FINAL = CAMINHO_PASTA_LOCAL / "Relatorios_OPEX.xlsx"
PASTA_LOGS = CAMINHO_PASTA_LOCAL / "Logs"
PASTA_LOGS.mkdir(exist_ok=True)

# Configuração de Log
logging.basicConfig(
    filename=PASTA_LOGS / f"log_execucao_{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.txt",
    filemode='w', 
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    datefmt='%d/%m/%Y %H:%M:%S'
)

# --- LISTA NEGRA DE REMETENTES ---
REMETENTES_IGNORAR = [
    "gleiciane.fragoso@grupofleury.com.br",
]

# Mapa de Meses
MAPA_MESES_INFO = {
    'january': ('Janeiro', '01'), 'jan': ('Janeiro', '01'), 'janeiro': ('Janeiro', '01'), '01': ('Janeiro', '01'),
    'february': ('Fevereiro', '02'), 'feb': ('Fevereiro', '02'), 'fevereiro': ('Fevereiro', '02'), '02': ('Fevereiro', '02'),
    'march': ('Março', '03'), 'mar': ('Março', '03'), 'março': ('Março', '03'), '03': ('Março', '03'),
    'april': ('Abril', '04'), 'apr': ('Abril', '04'), 'abril': ('Abril', '04'), '04': ('Abril', '04'),
    'may': ('Maio', '05'), 'maio': ('Maio', '05'), '05': ('Maio', '05'),
    'june': ('Junho', '06'), 'jun': ('Junho', '06'), 'junho': ('Junho', '06'), '06': ('Junho', '06'),
    'july': ('Julho', '07'), 'jul': ('Julho', '07'), 'julho': ('Julho', '07'), '07': ('Julho', '07'),
    'august': ('Agosto', '08'), 'aug': ('Agosto', '08'), 'agosto': ('Agosto', '08'), '08': ('Agosto', '08'),
    'september': ('Setembro', '09'), 'sep': ('Setembro', '09'), 'setembro': ('Setembro', '09'), 'set': ('Setembro', '09'), '09': ('Setembro', '09'),
    'october': ('Outubro', '10'), 'oct': ('Outubro', '10'), 'outubro': ('Outubro', '10'), 'out': ('Outubro', '10'), '10': ('Outubro', '10'),
    'november': ('Novembro', '11'), 'nov': ('Novembro', '11'), 'novembro': ('Novembro', '11'), '11': ('Novembro', '11'),
    'december': ('Dezembro', '12'), 'dec': ('Dezembro', '12'), 'dezembro': ('Dezembro', '12'), 'dez': ('Dezembro', '12'), '12': ('Dezembro', '12')
}

# Colunas padrão para garantir ordem
COLUNAS_PADRAO = [
    'Mes_Referencia', 'Data_Email', 'Remetente', 
    'Assunto_Email', 'Categoria_OPEX', 'Data_Processamento'
] 

# --- CONFIGURAÇÃO DOS FORNECEDORES ---
CONFIG_PADRAO = [
    {
        "Fornecedor": "Selbetti",
        "Palavras_Chave": "Faturamento Selbetti, Relatório Selbetti, RE: Faturamento Selbetti",
        "Nome_Aba": "Selbetti",
        "Categoria_OPEX": "Impressoras/Impressão"
    },
    {
        "Fornecedor": "Daycoval",
        "Palavras_Chave": "Faturamento Banco Daycoval, Relatório Daycoval",
        "Nome_Aba": "Daycoval",
        "Categoria_OPEX": "DAYCOVAL LEASING TI"
    },
    {
        "Fornecedor": "Positivo",
        "Palavras_Chave": "Faturamento Positivo, Locação",
        "Nome_Aba": "Positivo",
        "Categoria_OPEX": "POSITIVO LEASING TI"
    }
]

# --- FUNÇÕES AUXILIARES ---
                                                                 

def is_file_open(path):
    if not path.exists(): return False
    try:
        os.rename(path, path)
        return False
    except OSError:
        return True

def eh_coluna_financeira(nome_coluna):
    """
    Central de Controle: Define quais colunas são dinheiro.
    """
    nome_coluna = str(nome_coluna).lower()
    TERMOS_MOEDA = [
        'valor', 'total', 'liquido', 'bruto', 'realizado', 
        'vlr', 'custo', 'taxa', 'imposto', 'montante', 'r$'
    ]
    return any(termo in nome_coluna for termo in TERMOS_MOEDA)

def obter_email_remetente(msg):
    try:
        sender = msg.SenderEmailAddress
        if not sender: return ""
        if "O=" in sender and "@" not in sender:
            try:
                property_accessor = msg.Sender.PropertyAccessor
                sender = property_accessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x39FE001E")
            except: pass
        return sender.lower()
    except: return ""

def extrair_info_data_inteligente(assunto, corpo):
    padrao = r'(janeiro|jan|01|fevereiro|feb|02|março|mar|03|abril|apr|04|maio|may|05|junho|jun|06|julho|jul|07|agosto|aug|08|setembro|sep|set|09|outubro|oct|out|10|novembro|nov|11|dezembro|dec|dez|12)(?:/?\s?(?:de)?\s?(\d{2,4}))?'
    match = re.search(padrao, assunto.lower())
    if not match: match = re.search(padrao, corpo.lower())
    ano_atual = str(datetime.now().year)
    if match:
        mes_raw = match.group(1)
        ano_raw = match.group(2)
        mes_pt, mes_num = MAPA_MESES_INFO.get(mes_raw, (mes_raw.capitalize(), '00'))
        if ano_raw:
            if len(ano_raw) == 4 and ano_raw.startswith("20"): ano_final = ano_raw
            elif len(ano_raw) == 2: ano_final = "20" + ano_raw
            else: ano_final = ano_atual
        else: ano_final = ano_atual
        return mes_pt, mes_num, ano_final
    return None, None, ano_atual

def extrair_data_da_tabela(df):
    termos_data = ['data emissão', 'data emissao', 'dt emissao', 'data frs', 'dt frs', 'referencia', 'competencia', 'data de emissão', 'emissao']
    coluna_data_encontrada = None
    for col in df.columns:
        col_str = str(col).lower()
        if any(termo in col_str for termo in termos_data):
            coluna_data_encontrada = col
            break
    if not coluna_data_encontrada: return None, None, None
    try:
        valores = df[coluna_data_encontrada].dropna()
        if valores.empty: return None, None, None
        valor_data = valores.iloc[0]
        if isinstance(valor_data, str):
            try: dt = pd.to_datetime(valor_data, dayfirst=True)
            except: return extrair_info_data_inteligente(valor_data, "")
        elif isinstance(valor_data, (datetime, pd.Timestamp)): dt = valor_data
        else: return None, None, None
        
        meses_lista = ['Indef', 'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 
                       'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        mes_pt = meses_lista[dt.month] if 1 <= dt.month <= 12 else "Indefinido"
        return mes_pt, f"{dt.month:02d}", str(dt.year)
    except Exception as e:
        return None, None, None

def limpar_valor_monetario(valor):
    if isinstance(valor, (int, float)): return valor
    if isinstance(valor, str):
        limpo = valor.replace('R$', '').replace(' ', '').replace('.', '').replace(',', '.')
        try: return float(limpo)
        except: return 0.0
    return 0.0

def encontrar_cabecalho_correto(df, colunas_esperadas):
    """
    Tenta encontrar o cabeçalho real nas 3 primeiras linhas do DataFrame.
    """
    if not colunas_esperadas:
        # Fallback Positivo: procura tabela útil
        termos_validos = ['valor', 'total', 'liquido', 'bruto', 'vlr', 'data', 'emissao', 'nota', 'nf']
        
        cols_atuais = [str(c).lower() for c in df.columns]
        if any(t in c for c in cols_atuais for t in termos_validos):
            return df
            
        for i in range(min(3, len(df))):
            nova_header = df.iloc[i]
            str_header = [str(x).lower() for x in nova_header]
            if any(t in c for c in str_header for t in termos_validos):
                df_novo = df[i+1:].copy()
                df_novo.columns = nova_header
                return df_novo
        
        return None

    # Lógica normal
    cols_esperadas_norm = [c.lower().replace('.', '').replace(' ', '').strip() for c in colunas_esperadas]
    cols_atuais = [str(c).lower().replace('.', '').replace(' ', '').strip() for c in df.columns]
    if any(col in cols_atuais for col in cols_esperadas_norm): return df
    for i in range(min(3, len(df))):
        nova_header = df.iloc[i]
        df_novo = df[i+1:].copy()
        df_novo.columns = nova_header
        cols_teste = [str(c).lower().replace('.', '').replace(' ', '').strip() for c in df_novo.columns]
        if any(col in cols_teste for col in cols_esperadas_norm):
            return df_novo
    return None

def tratar_dataframe(df, config, metadados):
    if "colunas_renomear" in config: df = df.rename(columns=config["colunas_renomear"])
    df['Data_Email'] = pd.to_datetime(metadados['data_recebimento'], utc=True).tz_localize(None)
    df['Remetente'] = metadados['remetente']
    df['Assunto_Email'] = metadados['assunto']
    df['Mes_Referencia'] = metadados['mes_nome_pt'] if metadados['mes_nome_pt'] else ""
    df['Categoria_OPEX'] = config['classificacao_opex']
    df['Data_Processamento'] = datetime.now()

    # Aplica limpeza apenas em colunas financeiras
    colunas_valor = [col for col in df.columns if eh_coluna_financeira(col)]
    for col in colunas_valor: 
        df[col] = df[col].apply(limpar_valor_monetario)

    mes_chave = metadados['mes_nome_pt'] if metadados['mes_nome_pt'] else "SEM_DATA"
    col_desc = next((c for c in df.columns if 'Descricao' in str(c) or 'Linha' in str(c)), df.columns[0])
    
    # Tenta achar coluna de valor para a chave, senão pega a segunda coluna
    col_val = df.columns[1] if len(df.columns) > 1 else df.columns[0]
    for c in df.columns:
        if eh_coluna_financeira(c):
            col_val = c
            break
    
    df['Chave_Negocio_Temp'] = mes_chave + metadados['ano_full'] + "_" + \
                               df[col_desc].astype(str) + "_" + \
                               df[col_val].astype(str)
    
    cols_dados = [c for c in df.columns if c not in COLUNAS_PADRAO and c != 'Chave_Negocio_Temp']
    cols_finais = ['Mes_Referencia'] + cols_dados + ['Data_Email', 'Remetente', 'Assunto_Email', 'Categoria_OPEX', 'Data_Processamento', 'Chave_Negocio_Temp']
    
    return df[cols_finais]

def salvar_com_append_preservando_formatacao(df_novos, caminho_arquivo, nome_aba):
    if df_novos.empty: return

    # --- ARQUIVO NOVO ---
    if not caminho_arquivo.exists():
        if 'Chave_Negocio_Temp' in df_novos.columns:
            df_novos = df_novos.drop(columns=['Chave_Negocio_Temp'])
        
        with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
            df_novos.to_excel(writer, sheet_name=nome_aba, index=False)
            aplicar_estilo_inicial(writer.sheets[nome_aba], nome_aba)
        return

    try:
        # --- ARQUIVO EXISTENTE ---
        wb_check = load_workbook(caminho_arquivo, read_only=True)
        aba_existe = nome_aba in wb_check.sheetnames
        wb_check.close()

        # Cria aba se não existir
        if not aba_existe:
            if 'Chave_Negocio_Temp' in df_novos.columns:
                df_novos = df_novos.drop(columns=['Chave_Negocio_Temp'])
            
            with pd.ExcelWriter(caminho_arquivo, engine='openpyxl', mode='a') as writer:
                df_novos.to_excel(writer, sheet_name=nome_aba, index=False)
                aplicar_estilo_inicial(writer.sheets[nome_aba], nome_aba)
            return

        # Deduplicação
        df_existente = pd.read_excel(caminho_arquivo, sheet_name=nome_aba)
        colunas_no_excel = list(df_existente.columns)
        
        try:
            col_desc_ex = next((c for c in df_existente.columns if 'Descricao' in str(c) or 'Linha' in str(c)), df_existente.columns[1])
            
            # Busca chave de valor de forma inteligente
            col_val_ex = df_existente.columns[2]
            for c in df_existente.columns:
                if eh_coluna_financeira(c):
                    col_val_ex = c
                    break
            
            if 'Data_Email' in df_existente.columns:
                anos = pd.to_datetime(df_existente['Data_Email']).dt.year.astype(str)
            else:
                anos = str(datetime.now().year)

            chaves_existentes = set(
                df_existente['Mes_Referencia'].astype(str) + anos + "_" + 
                df_existente[col_desc_ex].astype(str) + "_" + 
                df_existente[col_val_ex].astype(str)
            )
        except:
            chaves_existentes = set()

        df_para_salvar = df_novos[~df_novos['Chave_Negocio_Temp'].isin(chaves_existentes)]
        
        if df_para_salvar.empty:
            logging.info(f"[{nome_aba}] Dados já existem. Nada a salvar.")
            return

        df_para_salvar = df_para_salvar.drop(columns=['Chave_Negocio_Temp'])

        # Alinhamento de colunas
        for col in colunas_no_excel:
            if col not in df_para_salvar.columns:
                df_para_salvar[col] = "" 
        
        df_para_salvar = df_para_salvar[colunas_no_excel]

        # --- GRAVAÇÃO COM FORMATAÇÃO ---
        wb = load_workbook(caminho_arquivo)
        ws = wb[nome_aba]
        
        # Mapeia quais índices são colunas financeiras (Baseado no cabeçalho do Excel)
        indices_financeiros = []
        for idx, col_name in enumerate(colunas_no_excel, start=1):
            if eh_coluna_financeira(col_name):
                indices_financeiros.append(idx)

        for r in dataframe_to_rows(df_para_salvar, index=False, header=False):
            ws.append(r)
            current_row = ws.max_row
            
            # Formatação Linha a Linha
            for col_idx, cell in enumerate(ws[current_row], start=1):
                cell.border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
                
                # APLICA FORMATAÇÃO VISUAL SE FOR COLUNA FINANCEIRA E VALOR NUMÉRICO
                if col_idx in indices_financeiros and isinstance(cell.value, (int, float)):
                     cell.number_format = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'

        # Atualiza Tabela
        nome_tabela_limpo = f"TB_{nome_aba.replace(' ', '_')}"
        if nome_tabela_limpo in ws.tables:
            tbl = ws.tables[nome_tabela_limpo]
            nova_ref = f"A1:{get_column_letter(ws.max_column)}{ws.max_row}"
            tbl.ref = nova_ref

        wb.save(caminho_arquivo)
        logging.info(f"[{nome_aba}] {len(df_para_salvar)} linhas anexadas com sucesso.")

    except Exception as e:
        logging.error(f"Erro ao anexar na aba {nome_aba}: {e}")

def aplicar_estilo_inicial(ws, sheet_name):
    max_col = ws.max_column
    letra_ultima_coluna = get_column_letter(max_col)
    
    fill_azul = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
    fill_destaque = PatternFill(start_color="ED7D31", end_color="ED7D31", fill_type="solid")
    header_font = Font(name="Calibri", size=11, color="FFFFFF", bold=True)
    
    # Identifica financeiros
    indices_financeiros = []

    for idx, cell in enumerate(ws[1], start=1):
        col_name = str(cell.value)
        
        if eh_coluna_financeira(col_name):
            indices_financeiros.append(idx)

        if col_name in COLUNAS_PADRAO:
            cell.fill = fill_destaque
        else:
            cell.fill = fill_azul
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center')

    # Aplica nas linhas iniciais (se houver dados)
    for row in ws.iter_rows(min_row=2, max_col=max_col):
        for col_idx, cell in enumerate(row, start=1):
            if col_idx in indices_financeiros and isinstance(cell.value, (int, float)):
                cell.number_format = '_-R$ * #,##0.00_-;-R$ * #,##0.00_-;_-R$ * "-"??_-;_-@_-'

    ref_tabela = f"A1:{letra_ultima_coluna}{ws.max_row}"
    nome_tabela_limpo = f"TB_{sheet_name.replace(' ', '_')}"
    tabela = Table(displayName=nome_tabela_limpo, ref=ref_tabela)
    estilo = TableStyleInfo(name="TableStyleLight1", showFirstColumn=False, showLastColumn=False, showRowStripes=False, showColumnStripes=False)
    tabela.tableStyleInfo = estilo
    ws.add_table(tabela)

# --- PROCESSO PRINCIPAL ---

def realizar_backup_seguranca():
    """
    Cria uma cópia do arquivo Excel atual antes de qualquer alteração.
    O arquivo é salvo na pasta 'Backups' com data e hora no nome.
    """
    if not ARQUIVO_FINAL.exists():
        logging.info("Arquivo final não existe. Pulinho backup.")
        print("Arquivo final não existe. Pulinho backup.")
        return

    try:
        # 1. Define onde será a pasta de backups
        pasta_backup = CAMINHO_PASTA_LOCAL / "Backups"
        
        # 2. Cria a pasta se ela não existir
        pasta_backup.mkdir(exist_ok=True)
        
        # 3. Define o nome do arquivo com DATA e HORA (para não substituir o anterior)
        timestamp = datetime.now().strftime('%Y-%m-%d_%H-%M-%S')
        nome_backup = f"Backup_{timestamp}_Relatorios_OPEX.xlsx"
        caminho_backup = pasta_backup / nome_backup
        
        # 4. Copia o arquivo
        shutil.copy2(ARQUIVO_FINAL, caminho_backup)
        
        print(f" -> BACKUP REALIZADO: {nome_backup}")
        logging.info(f"Backup de segurança criado: {nome_backup}")
        
    except Exception as e:
        print(f"ERRO AO FAZER BACKUP: {e}")
        logging.error(f"Falha no backup: {e}")    

def inicializar_aba_config():
    """
    Verifica se a aba 'config_fornecedor' existe. 
    Se não existir, cria ela, coloca o cabeçalho e os dados padrão.
    """
    try:
        wb = None
        salvar_necessario = False
        NOME_ABA_CONFIG = "config_fornecedor" # <--- MUDAMOS O NOME AQUI

        # 1. Abre ou Cria o Arquivo
        if not ARQUIVO_FINAL.exists():
            print(f"--- CRIANDO ARQUIVO NOVO EM: {ARQUIVO_FINAL} ---")
            wb = Workbook() 
            if 'Sheet' in wb.sheetnames: del wb['Sheet']
            salvar_necessario = True
        else:
            wb = load_workbook(ARQUIVO_FINAL)

        # 2. Verifica/Cria a aba específica
        if NOME_ABA_CONFIG not in wb.sheetnames:
            print(f"Criando aba '{NOME_ABA_CONFIG}'...")
            ws = wb.create_sheet(NOME_ABA_CONFIG, 0) # Cria na primeira posição

            # --- ESCREVE O CABEÇALHO ---
            cabecalho = ["Fornecedor", "Palavras_Chave", "Nome_Aba", "Categoria_OPEX"]
            ws.append(cabecalho)

            # --- PREENCHE DADOS PADRÃO ---
            # Garante que CONFIG_PADRAO seja tratado como lista
            dados = CONFIG_PADRAO
            if isinstance(CONFIG_PADRAO, dict): dados = list(CONFIG_PADRAO.values())

            for item in dados:
                if isinstance(item, dict):
                    ws.append([
                        item.get("Fornecedor", ""),
                        item.get("Palavras_Chave", ""),
                        item.get("Nome_Aba", ""),
                        item.get("Categoria_OPEX", "")
                    ])

            # --- FORMATAÇÃO (Para não confundir o usuário) ---
            # Pinta o cabeçalho de Preto com letra Branca
            for cell in ws[1]:
                cell.font = Font(bold=True, color="FFFFFF")
                cell.fill = PatternFill(start_color="000000", end_color="000000", fill_type="solid")
                cell.alignment = Alignment(horizontal='center')
            
            # Ajusta largura das colunas para ficar legível
            ws.column_dimensions['A'].width = 15 # Fornecedor
            ws.column_dimensions['B'].width = 40 # Palavras Chave
            ws.column_dimensions['C'].width = 15 # Nome Aba
            ws.column_dimensions['D'].width = 25 # Categoria
            
            salvar_necessario = True

        if salvar_necessario:
            wb.save(ARQUIVO_FINAL)
            print(f" -> ARQUIVO SALVO COM ABA '{NOME_ABA_CONFIG}'.")
        
        wb.close()
        
    except Exception as e:
        if "Permission denied" in str(e):
            print("ERRO CRÍTICO: O Excel está aberto. Feche-o para criar a configuração.")
        logging.error(f"Erro ao inicializar aba config: {e}")

def carregar_configuracoes_do_excel():
    """Lê a aba 'config_fornecedor' e transforma no dicionário."""
    
    inicializar_aba_config()
    
    config_dict = {} 
    NOME_ABA_CONFIG = "config_fornecedor" # <--- MUDAMOS O NOME AQUI TAMBÉM

    try:
        if not ARQUIVO_FINAL.exists():
             raise Exception("Arquivo não encontrado.")

        # Lê a aba correta
        df_config = pd.read_excel(ARQUIVO_FINAL, sheet_name=NOME_ABA_CONFIG)
        df_config = df_config.dropna(how='all') 

        for _, row in df_config.iterrows():
            fornecedor = str(row["Fornecedor"])
            if fornecedor == 'nan': continue

            palavras = str(row["Palavras_Chave"]).split(',')
            palavras_limpas = [p.strip() for p in palavras if p.strip()]
            
            config_dict[fornecedor] = {
                "assuntos_possiveis": palavras_limpas,
                "nome_aba": str(row["Nome_Aba"]),
                "classificacao_opex": str(row["Categoria_OPEX"]),
                "colunas_renomear": {}
            }
        
        print(f" -> Configuração carregada: {len(config_dict)} fornecedores.")
        return config_dict

    except Exception as e:
        logging.error(f"Erro lendo Excel ({e}). Usando padrão de memória.")
        print(f"Aviso: Usando padrão de memória (Erro: {e})")
        
        # Fallback
        fallback_dict = {}
        dados = CONFIG_PADRAO
        if isinstance(CONFIG_PADRAO, dict): dados = list(CONFIG_PADRAO.values())

        for item in dados:
            if isinstance(item, dict):
                fallback_dict[item["Fornecedor"]] = {
                    "assuntos_possiveis": item["Palavras_Chave"].split(','),
                    "nome_aba": item["Nome_Aba"],
                    "classificacao_opex": item["Categoria_OPEX"],
                    "colunas_renomear": {}
                }
        return fallback_dict

def enviar_email_resumo(resumo_dados):
    """
    Gera o e-mail e MOSTRA NA TELA (.Display) para conferência.
    """

    if not resumo_dados:
        print(" -> Nada processado. E-mail de resumo não será gerado.")
        logging.info("Nada processado. E-mail ignorado.")
        return
    
    try:
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)

        # SEU E-MAIL AQUI
        mail.To = "matheus.augusto@grupofleury.com.br; lucas.rosa@grupofleury.com.br"
        mail.CC = "gustavo.guanas@grupofleury.com.br; jeferson.costa@grupofleury.com.br"

        data_hoje = datetime.now().strftime('%d/%m/%Y')
        mail.Subject = f"Resumo Processamento OPEX - {data_hoje}"

        #--- MONTAGEM DO HTML ---
        html = f"""
        <html>
        <body style="font-family: Calibri, sans-serif;">
            <h2 style="color: #2F75B5;">Processamento Finalizado com Sucesso</h2>
            <p>O robô Sentinela OPEX finalizou a execução de hoje ({data_hoje}).</p>
            
            <table style="border-collapse: collapse; width: 100%; max-width: 600px;">
                <tr style="background-color: #f2f2f2;">
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: left;">Fornecedor</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: center;">Qtd E-mails</th>
                    <th style="border: 1px solid #ddd; padding: 8px; text-align: right;">Valor Total</th>
                </tr>
        """

        total_geral_valor = 0
        total_geral_qtd = 0

        for forn, dados in resumo_dados.items():
            qtd = dados.get('qtd', 0)
            valor = dados.get('valor', 0.0)
            total_geral_valor += valor
            total_geral_qtd += qtd

            valor_fmt = f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
            
            html += f"""
                <tr>
                    <td style="border: 1px solid #ddd; padding: 8px;">{forn}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{qtd}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">{valor_fmt}</td>
                </tr>
            """

        total_fmt = f"R$ {total_geral_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        html += f"""
                <tr style="font-weight: bold; background-color: #e6f3ff;">
                    <td style="border: 1px solid #ddd; padding: 8px;">TOTAL GERAL</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: center;">{total_geral_qtd}</td>
                    <td style="border: 1px solid #ddd; padding: 8px; text-align: right;">{total_fmt}</td>
                </tr>
            </table>
            <br>
            <p style="font-size: 11px; color: #888;">Este e-mail foi gerado automaticamente pelo Robô Sentinela OPEX.</p>
        </body>
        </html>
        """

        mail.HTMLBody = html
        
        # --- MUDANÇA CRUCIAL AQUI ---
        # mail.Send()  <-- Comentado para evitar bloqueio silencioso
        mail.Display() # <-- VAI ABRIR A JANELA DO E-MAIL NA SUA TELA
        
        print("\n -> JANELA DE E-MAIL ABERTA! VERIFIQUE E CLIQUE EM ENVIAR.")
        logging.info("Janela de e-mail de resumo aberta.")

    except Exception as e:
        print(f"Erro ao gerar e-mail de resumo: {e}")
        logging.error(f"Erro no e-mail: {e}")

def executar_pipeline():
    print("\n--- INICIANDO PROCESSAMENTO OPEX ---")
    logging.info("--- Iniciando Processamento ---")

    if is_file_open(ARQUIVO_FINAL):
        logging.critical(f"O arquivo {ARQUIVO_FINAL} está ABERTO. Feche-o.")
        print("ERRO CRÍTICO: O arquivo Excel está aberto. Por favor, feche-o.")
        return
    
    print("\n--- INICIANDO BACKUP DE SEGURANÇA ---")
    realizar_backup_seguranca()

    # --- Variável para guardar os dados do resumo ---
    stats_geral = {} 
    # -----------------------------------------------------

    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
        inbox = outlook.GetDefaultFolder(6)
        msgs_para_mover_com_destino = []

        print(" -> Carregando configurações...")
        config_atual = carregar_configuracoes_do_excel()

        # Proteção extra caso venha como lista
        if isinstance(config_atual, list):
            logging.warning("Configuração veio como lista. Convertendo...")
            novo_dict = {}
            for item in config_atual:
                if isinstance(item, dict) and "Fornecedor" in item: 
                     palavras = item.get("Palavras_Chave", "")
                     lista_palavras = palavras.split(',') if isinstance(palavras, str) else []
                     novo_dict[item["Fornecedor"]] = {
                        "assuntos_possiveis": lista_palavras,
                        "nome_aba": item.get("Nome_Aba", ""),
                        "classificacao_opex": item.get("Categoria_OPEX", ""),
                        "colunas_renomear": {}
                      }
            config_atual = novo_dict

        if not config_atual:
            print("ERRO: Nenhuma configuração carregada.")
            return

        for fornecedor, config in config_atual.items():
            print(f"\nVerificando fornecedor: {fornecedor}")
            logging.info(f"Verificando e-mails de: {fornecedor}")
            mensagens = inbox.Items
            mensagens.Sort("[ReceivedTime]", True)
            assuntos_alvo = [a.lower() for a in config.get("assuntos_possiveis", [])]
            dfs_novos = []

            for msg in mensagens:
                try:
                    if getattr(msg, 'Class', 0) != 43: continue
                    try:
                        if msg.ReceivedTime.replace(tzinfo=None) < DATA_INICIO_LEITURA: continue
                    except: continue

                    assunto_msg = msg.Subject.lower()
                    if not any(alvo in assunto_msg for alvo in assuntos_alvo): continue 
                    
                    remetente_email = obter_email_remetente(msg)
                    # --- ATENÇÃO: COMENTE A LINHA ABAIXO PARA TESTAR COM SEU E-MAIL ---
                    # if remetente_email in REMETENTES_IGNORAR: 
                    #     print(f" -> Ignorado: Remetente bloqueado ({remetente_email})")
                    #     continue
                    # ------------------------------------------------------------------

                    print(f" -> PROCESSANDO: {msg.Subject}")
                    logging.info(f"Processando e-mail: {msg.Subject}")
                    html = msg.HTMLBody
                    todas_tabelas = pd.read_html(StringIO(html), header=0)
                    if not todas_tabelas:
                        todas_tabelas = pd.read_html(StringIO(html), header=None)
                        if not todas_tabelas: continue

                    tabela_alvo = None
                    colunas_esperadas = list(config['colunas_renomear'].keys())
                    for tb in todas_tabelas:
                        tb_ajustada = encontrar_cabecalho_correto(tb, colunas_esperadas)
                        if tb_ajustada is not None:
                            tabela_alvo = tb_ajustada
                            break
                    
                    if tabela_alvo is None and fornecedor == "Positivo":
                        for tb in todas_tabelas:
                            if len(tb.columns) > 1:
                                cols_texto = [str(c).lower() for c in tb.columns]
                                if any(termo in col for col in cols_texto for termo in ['valor', 'total', 'liq', 'bruto', 'r$']):
                                    tabela_alvo = tb
                                    logging.info("Tabela financeira encontrada (Fallback Positivo).")
                                    break

                    if tabela_alvo is None:
                        logging.warning(f"Tabela correta não encontrada em: {msg.Subject}.")
                        continue
                    
                    mes_pt, mes_num, ano_full = extrair_data_da_tabela(tabela_alvo)
                    if not mes_pt: mes_pt, mes_num, ano_full = extrair_info_data_inteligente(msg.Subject, msg.Body)
                    
                    metadados = {
                        'data_recebimento': str(msg.ReceivedTime),
                        'remetente': msg.SenderName,
                        'assunto': msg.Subject,
                        'mes_nome_pt': mes_pt, 
                        'ano_full': ano_full   
                    }

                    df_tratado = tratar_dataframe(tabela_alvo, config, metadados)
                    dfs_novos.append(df_tratado)
                    
                    pasta_ano = f"Processados_OPEX {ano_full}"
                    pasta_mes = f"{mes_num} - {mes_pt}" if mes_pt else "00 - A Classificar"
                    msgs_para_mover_com_destino.append((msg, pasta_ano, pasta_mes, fornecedor))

                except Exception as e:
                    logging.error(f"Erro no loop de msg: {e}")
                    continue
            
            if dfs_novos:
                full_new_data = pd.concat(dfs_novos)
                salvar_com_append_preservando_formatacao(full_new_data, ARQUIVO_FINAL, config['nome_aba'])
                
                # --- COLETA DADOS PARA O E-MAIL ---
                qtd = len(dfs_novos)
                soma_valor = 0.0
                for col in full_new_data.columns:
                    if eh_coluna_financeira(col):
                        soma_valor = full_new_data[col].sum()
                        break
                
                stats_geral[fornecedor] = {'qtd': qtd, 'valor': soma_valor}
                # ----------------------------------------
            else:
                logging.info(f"Nenhum dado novo para {fornecedor}.")

        print("\nMovendo e-mails processados...")

        for msg, pasta_ano_nome, pasta_mes_nome, pasta_fornecedor_nome in msgs_para_mover_com_destino:
            try:
                msg.UnRead = False 
                try: pasta_ano = inbox.Folders(pasta_ano_nome)
                except: pasta_ano = inbox.Folders.Add(pasta_ano_nome)
                try: pasta_mes = pasta_ano.Folders(pasta_mes_nome)
                except: pasta_mes = pasta_ano.Folders.Add(pasta_mes_nome)
                try: pasta_final = pasta_mes.Folders(pasta_fornecedor_nome)
                except: pasta_final = pasta_mes.Folders.Add(pasta_fornecedor_nome)
                msg.Move(pasta_final)
            except Exception as e: logging.error(f"Erro ao mover email: {e}")
        
        # --- DISPARA O E-MAIL NO FINAL ---
        print("\nGerando relatório e enviando e-mail...")
        enviar_email_resumo(stats_geral)
        # ---------------------------------------

        print("\n--- PROCESSO CONCLUÍDO COM SUCESSO ---")

    except Exception as e:
        logging.critical(f"Falha critica na execução: {e}")
        print(f"Erro Crítico: {e}")

if __name__ == "__main__":
    executar_pipeline()