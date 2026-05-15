#!/usr/bin/env python
# coding: utf-8

# In[ ]:


from datetime import datetime
from flask import Flask, flash, request, redirect, url_for, session, render_template, jsonify
from itertools import islice
from openpyxl import load_workbook
from pathlib import Path

import json
import openpyxl
import openpyxl.workbook
import os
import pandas as pd
import pythoncom
import re
import html
import sys
import win32com.client as win32

app = Flask(__name__, template_folder=r'C:\Users\usuario\desktop\projeto\Templates')
app.secret_key = os.urandom(24)

def resource_path(relative_path):
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

UPLOAD_FOLDER = r'C:\Users\usuario\desktop\projeto'
ALLOWED_EXTENSIONS = {'xlsm','xlsx','xls','csv'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.',1)[1].lower() in ALLOWED_EXTENSIONS

texto_email = None
texto_email_tratado = None

def df_to_outlook_table(df,
                        header_bg="#004b2c",
                        header_color="#ffffff",
                        zebra_even_bg="#f3f3f3",
                        border_color="#dddddd",
                        font_family="Arial, sans-serif",
                        font_size="12px",
                        cell_padding="8px 10px",
                        table_min_width="200px",
                        outer_margin="15px 0",
                        column_width_px="auto",
                        active_row_selector=None):
    table_style = (
        f"border-collapse: collapse;"
        f"margin: {outer_margin};"
        f"font-size: {font_size};"
        f"font-family: {font_family};"
        f"min-width: {table_min_width};"
        f"width: 100%;"
    )

    thead_tr_style = (
        f"background-color: {header_bg};"
        f"color: {header_color};"
        f"text-align: left;"
        f"border-bottom: 2px solid {header_bg};"
    )

    th_td_base_style = (
        f"padding: {cell_padding};"
        f"border-bottom: 1px solid {border_color};"
        f"mso-line-height-rule: exactly;"
        f"text-align: left;"
    )
   
    if column_width_px != "auto":
        th_td_base_style += f"width: {column_width_px};"

    columns = list(df.columns)
    thead_cells = "".join(
        f'<th style="{th_td_base_style}">{html.escape(str(col))}</th>'
        for col in columns
    )
    thead = f'<thead><tr style="{thead_tr_style}">{thead_cells}</td></thead>'

    tbody_rows = []
    for i, (_, row) in enumerate(df.iterrows()):
        row_bg = zebra_even_bg if (i % 2 == 1) else "transparent"
        tr_style_extra = f"background-color: {row_bg};"

        is_active = False
        if callable(active_row_selector):
            row_dict = {col: row[col] for col in columns}
            is_active = bool(active_row_selector(row_dict))
            if is_active:
                tr_style_extra += "font-weight: bold; color: #009879;"

        tds = []
        for col in columns:
            val = row[col]
            if val is None or (isinstance(val, float) and str(val) == "nan"):
                val_str = ""
            else:
                val_str = html.escape(str(val))
            tds.append(f'<td style="{th_td_base_style}">{val_str}</td>')

        tbody_rows.append(f'<tr style="{tr_style_extra}">' + "".join(tds) + "</tr>")

    tbody = "<tbody>" + "".join(tbody_rows) + "</tbody>"
    table_html = f'<table style="{table_style}">{thead}{tbody}</table>'
    return table_html

def processar_calculos(texto_html, df):
    def calcular_expressao(match):
        expr = match.group(0)
        
        padrao_func = r'(sum|mean|count|max|min)\[([^\]]+)\]'
        match_func = re.search(padrao_func, expr)
        
        if not match_func:
            return expr
        
        func = match_func.group(1)
        coluna = match_func.group(2).strip()
        
        if coluna not in df.columns:
            return f"[Coluna '{coluna}' não encontrada]"

        try:
            if func == 'sum':
                valor = df[coluna].sum()
            elif func == 'mean':
                valor = df[coluna].mean()
            elif func == 'count':
                valor = df[coluna].count()
            elif func == 'max':
                valor = df[coluna].max()
            elif func == 'min':
                valor = df[coluna].min()
            else:
                return expr
            
            if isinstance(valor, (int, float)):
                if isinstance(valor, float) and valor.is_integer():
                    return str(int(valor))
                elif isinstance(valor, float):
                    return f"{valor:.2f}".replace('.', ',')
                else:
                    return str(valor)
            return str(valor)
            
        except Exception as e:
            return f"[Erro: {e}]"
    
    padrao = r'\{?(sum|mean|count|max|min)\[[^\]]+\]\}?'
    
    return re.sub(padrao, calcular_expressao, texto_html)

def get_logged_user_email():
    try:
        pythoncom.CoInitialize()
        outlook = win32.Dispatch('Outlook.Application')
        session = outlook.Session
        user = session.CurrentUser

        # Obter EntryID e acessar o objeto AddressEntry
        address_entry = user.AddressEntry

        # Verifica se é Exchange
        if address_entry.Type == "EX":
            exchange_user = address_entry.GetExchangeUser()
            if exchange_user:
                return exchange_user.PrimarySmtpAddress
        else:
            return user.Address

    except Exception as e:
        print(f"Erro ao obter e-mail do usuário: {e}")
        return None

def normalizar_caminho(caminho):
    """Normaliza o caminho do arquivo, convertendo barras invertidas para barras normais"""
    if not caminho:
        return ""
    # Substitui barras invertidas por barras normais
    caminho_normalizado = caminho.replace('\\', '/')
    # Remove qualquer caractere de escape
    caminho_normalizado = caminho_normalizado.strip()
    return caminho_normalizado

@app.route('/')
def index():
    return render_template("index.html")


@app.route('/upload', methods=['GET', 'POST'])
def upload_func():
    global texto_email_tratado
    
    if request.method == 'GET':
        return render_template("upload.html", columns=[])
    
    if request.method == 'POST':

        texto_email_tratado = request.form.get('formattedText', '')
        
        data_envio = datetime.now().strftime('%d/%m/%Y %H:%M')
        
        caminho_arquivo = request.form.get('caminho_arquivo', '').replace('"', '')
        coluna_email = request.form.get('coluna_email', '')
        coluna_nome_destinatario = request.form.get('coluna_nome_destinatario', '')
        sheet = request.form.get('sheet_group', '')
        assunto_original = request.form.get('assunto', '')  # Renomeado para assunto_original
        opcao_teste = request.form.get('opcao_teste', '')
        data_source = request.form.get('data_source', '')
        selected_columns_json = request.form.get('selected_columns', '[]')
        
        try:
            selected_columns = json.loads(selected_columns_json)
        except Exception as e:
            print(f"Erro ao parsear selected_columns: {e}")
            selected_columns = []
        
        # ========== VERIFICAÇÃO DE DUPLICIDADE ANTES DE TUDO ==========
        colunas_salvas_temp = ','.join(selected_columns)
        
        registro_path = f'{UPLOAD_FOLDER}/Registro/Registro.xlsx'
        registro_duplicado = False
        
        if os.path.exists(registro_path):
            try:
                registro_df = pd.read_excel(registro_path, index_col=None)
                registro_df.columns = registro_df.columns.str.strip()
                
                for idx, row in registro_df.iterrows():
                    if (str(row.get('Caminho Arquivo', '')) == caminho_arquivo and
                        str(row.get('Folha', '')) == sheet and
                        str(row.get('Nome Coluna E-mail', '')) == coluna_email and
                        str(row.get('Nome Coluna Destinatário', '')) == coluna_nome_destinatario and
                        str(row.get('Colunas utilizadas', '')) == colunas_salvas_temp and
                        str(row.get('Assunto', '')) == assunto_original and
                        str(row.get('Corpo', '')) == texto_email_tratado and
                        str(row.get('Origem Dados', '')) == data_source):
                        
                        registro_duplicado = True
                        break
            except Exception as e:
                print(f"Erro ao ler registro para verificação: {e}")
        
        if registro_duplicado:
            flash('⚠️ Configuração já existe! Nenhum novo registro foi criado e nenhum e-mail foi enviado.', 'warning')
            return redirect(url_for('table_func'))
        # ========== FIM DA VERIFICAÇÃO ==========
        
        try:
            pythoncom.CoInitialize()
            outlook = win32.Dispatch('outlook.application')
        except Exception as e:
            print(f"Erro ao conectar ao Outlook: {e}")
            return jsonify({'success': False, 'error': f'Erro no Outlook: {str(e)}'})
        
        user_email = get_logged_user_email()
        
        pasta = None
        filename = None
        df_base = None
        df_email = None
        colunas_filtradas = []
        
        if data_source in ['file', 'both']:
            
            if 'file' not in request.files:
                print("ERRO: Nenhum arquivo encontrado na requisição")
                return jsonify({'success': False, 'error': 'Nenhum arquivo selecionado'})
            
            file = request.files['file']
            
            if file.filename == '':
                print("ERRO: Nome do arquivo vazio")
                return jsonify({'success': False, 'error': 'Nome do arquivo vazio'})
            
            if not allowed_file(file.filename):
                print(f"ERRO: Tipo de arquivo não permitido: {file.filename}")
                return jsonify({'success': False, 'error': 'Tipo de arquivo não permitido'})
            
            filename = file.filename
            filename_sem_extensao = filename.replace(".xlsx", "").replace(".xls", "").replace(".csv", "")
            
            pasta = f"{UPLOAD_FOLDER}/{filename_sem_extensao} {sheet}"
            if not os.path.isdir(pasta):
                os.makedirs(pasta)
            
            file.save(os.path.join(pasta, filename))

            if filename.endswith('.csv'):
                caminho_normalizado = Path(caminho_arquivo).resolve()
                df = pd.read_csv(caminho_normalizado)
            else:
                caminho_normalizado = Path(caminho_arquivo).resolve()
                df = pd.read_excel(caminho_normalizado, sheet_name=sheet)
            
            for col in selected_columns:
                if col in df.columns:
                    colunas_filtradas.append(col)

            colunas_lista = colunas_filtradas.copy()
            colunas_lista.append(coluna_email)
            colunas_lista.append(coluna_nome_destinatario)
            
            base = "base"
            
            if coluna_email in df.columns:
                df_sem_duplicatas = df.ffill().dropna(axis=0).drop_duplicates(subset=[coluna_email], keep='first')
            else:
                df_sem_duplicatas = df.ffill().dropna(axis=0)
            
            df_sem_duplicatas.to_excel(
                f"{pasta}/{base} - {filename}",
                index=False,
                columns=colunas_lista,
                sheet_name=sheet
            )
            
            df_base = pd.read_excel(f'{pasta}/{filename}', sheet_name=sheet)
            df_email = pd.read_excel(f'{pasta}/{base} - {filename}', sheet_name=sheet)
        
        elif data_source == 'dataframe':
            if 'file' not in request.files:
                flash('Nenhum arquivo selecionado para gerar o DataFrame', 'error')
                return redirect(url_for('table_func'))
            
            file = request.files['file']
            if file.filename == '':
                flash('Nome do arquivo vazio', 'error')
                return redirect(url_for('table_func'))
            
            filename = file.filename
            filename_sem_extensao = filename.replace(".xlsx", "").replace(".xls", "").replace(".csv", "")
            
            pasta = f"{UPLOAD_FOLDER}/{filename_sem_extensao} {sheet}"
            if not os.path.isdir(pasta):
                os.makedirs(pasta)
            
            file.save(os.path.join(pasta, filename))
            
            if filename.endswith('.csv'):
                df = pd.read_csv(os.path.join(pasta, filename))
            else:
                df = pd.read_excel(os.path.join(pasta, filename), sheet_name=sheet)
            
            for col in selected_columns:
                if col in df.columns:
                    colunas_filtradas.append(col)
            
            if coluna_email in df.columns:
                df_sem_duplicatas = df.ffill().dropna(axis=0).drop_duplicates(subset=[coluna_email], keep='first')
            else:
                df_sem_duplicatas = df.ffill().dropna(axis=0)
            
            df_base = df
            df_email = df_sem_duplicatas  # USA O SEM DUPLICATAS

        elif data_source == 'none':
            # Mesmo comportamento do dataframe, mas sem anexar arquivo e sem tabela
            if 'file' not in request.files:
                flash('Nenhum arquivo selecionado', 'error')
                return redirect(url_for('table_func'))
            
            file = request.files['file']
            if file.filename == '':
                flash('Nome do arquivo vazio', 'error')
                return redirect(url_for('table_func'))
            
            filename = file.filename
            filename_sem_extensao = filename.replace(".xlsx", "").replace(".xls", "").replace(".csv", "")
            
            pasta = f"{UPLOAD_FOLDER}/{filename_sem_extensao} {sheet}"
            if not os.path.isdir(pasta):
                os.makedirs(pasta)
            
            file.save(os.path.join(pasta, filename))
            
            if filename.endswith('.csv'):
                df = pd.read_csv(os.path.join(pasta, filename))
            else:
                df = pd.read_excel(os.path.join(pasta, filename), sheet_name=sheet)
            
            for col in selected_columns:
                if col in df.columns:
                    colunas_filtradas.append(col)
            
            # Remove duplicatas por email
            if coluna_email in df.columns:
                df_sem_duplicatas = df.ffill().dropna(axis=0).drop_duplicates(subset=[coluna_email], keep='first')
            else:
                df_sem_duplicatas = df.ffill().dropna(axis=0)
            
            df_base = df
            df_email = df_sem_duplicatas
            # pasta e filename mantidos, mas não serão anexados porque a condição de anexo verifica data_source in ['file', 'both']
            
            # Remove o placeholder {dataframe} do corpo do email
            texto_email_tratado = re.sub(r'\{dataframe\}', '', texto_email_tratado)
            
            print(f"Modo Nenhum ativado. Enviando para {len(df_email)} destinatários. Colunas salvas: {colunas_filtradas}")
        
        corpo_html = texto_email_tratado if texto_email_tratado else ""
        
        def enviar_email(destinatario, nome_destinatario, filtro_df, row_email=None, idx=0):
            try:
                # CORREÇÃO 1: Substituir {destinatario} no assunto
                assunto_processado = assunto_original.replace('{destinatario}', nome_destinatario or '')
                
                corpo_final = corpo_html

                if filtro_df is not None and not filtro_df.empty:
                    corpo_final = processar_calculos(corpo_final, filtro_df)
                else:
                    corpo_final = re.sub(r'\{[^}]+\}', '', corpo_final)
                
                if data_source in ['dataframe', 'both'] and filtro_df is not None and not filtro_df.empty:
                    html_table = df_to_outlook_table(
                        filtro_df,
                        header_bg="#004b2c",
                        header_color="#ffffff",
                        zebra_even_bg="#f3f3f3"
                    )
                    corpo_final = corpo_final.replace('{dataframe}', html_table)
                else:
                    corpo_final = corpo_final.replace('{dataframe}', '')
                
                nome = nome_destinatario.rstrip() if isinstance(nome_destinatario, str) else str(nome_destinatario) if nome_destinatario else ""
                corpo_final = corpo_final.replace('{destinatario}', nome)
                corpo_final = re.sub(r'\{[^}]+\}', '', corpo_final)
                
                mail = outlook.CreateItem(0)
                mail.Subject = assunto_processado  # CORREÇÃO: Usa assunto com placeholder substituído
                
                if opcao_teste == "Sim":
                    mail.To = user_email
                else:
                    mail.To = destinatario
                
                email_html = render_template('placeholders.html',
                    corpo_mensagem=corpo_final,
                    tabela_dataframe=''
                )
                mail.HTMLBody = email_html
                
                if data_source in ['file', 'both'] and row_email is not None and pasta and filename:
                    nome_arquivo = row_email[coluna_nome_destinatario] if coluna_nome_destinatario in row_email else nome
                    if isinstance(nome_arquivo, str):
                        nome_arquivo = nome_arquivo.rstrip()
                    else:
                        nome_arquivo = str(nome_arquivo) if nome_arquivo else "arquivo"
                    
                    nome_arquivo = re.sub(r'[\\/*?:"<>|]', '', nome_arquivo)
                    arquivo_anexo = f"{pasta}/{nome_arquivo} - {filename}"
                    
                    if not filtro_df.empty and colunas_filtradas:
                        filtro_df.to_excel(arquivo_anexo, index=False)
                    
                    if os.path.exists(arquivo_anexo):
                        mail.Attachments.Add(arquivo_anexo)
                
                mail.Send()
                return True
                
            except Exception as e:
                print(f"Erro ao enviar e-mail: {e}")
                return False
        
        if df_email is not None and not df_email.empty:
            if opcao_teste == "Sim":
                total = min(5, len(df_email))
                for idx, row_email in df_email.iloc[0:5].iterrows():
                    if df_base is not None and coluna_email in df_base.columns:
                        filtro = df_base[df_base[coluna_email] == str(row_email[coluna_email])]
                        filtro = filtro[colunas_filtradas] if not filtro.empty and colunas_filtradas else pd.DataFrame()
                    else:
                        filtro = pd.DataFrame()
                    
                    nome = row_email[coluna_nome_destinatario] if coluna_nome_destinatario in row_email else ""
                    
                    enviar_email(user_email, nome, filtro, row_email, idx)
            else:
                total = len(df_email)
                for idx, row_email in df_email.iterrows():
                    if df_base is not None and coluna_email in df_base.columns:
                        filtro = df_base[df_base[coluna_email] == str(row_email[coluna_email])]
                        filtro = filtro[colunas_filtradas] if not filtro.empty and colunas_filtradas else pd.DataFrame()
                    else:
                        filtro = pd.DataFrame()
                    
                    nome = row_email[coluna_nome_destinatario] if coluna_nome_destinatario in row_email else ""
                    destinatario = row_email[coluna_email] if coluna_email in row_email else ""
                    
                    if destinatario:
                        enviar_email(destinatario, nome, filtro, row_email, idx)
        
        # SALVA NO REGISTRO
        try:
            if os.path.exists(registro_path):
                wb = load_workbook(registro_path, data_only=True)
                if 'Sheet1' in wb.sheetnames:
                    pag_reg = wb['Sheet1']
                else:
                    pag_reg = wb.active
                    pag_reg.title = 'Sheet1'
            else:
                wb = openpyxl.Workbook()
                pag_reg = wb.active
                pag_reg.title = 'Sheet1'
            
            colunas_salvas = ','.join(colunas_filtradas)
            
            linha_encontrada = False
            for row in range(2, pag_reg.max_row + 2):
                if pag_reg.cell(row=row, column=1).value is None:
                    pag_reg.cell(row=row, column=1, value=caminho_arquivo)
                    pag_reg.cell(row=row, column=2, value=sheet)
                    pag_reg.cell(row=row, column=3, value=coluna_email)
                    pag_reg.cell(row=row, column=4, value=coluna_nome_destinatario)
                    pag_reg.cell(row=row, column=5, value=colunas_salvas)
                    pag_reg.cell(row=row, column=6, value=assunto_original)
                    pag_reg.cell(row=row, column=7, value=texto_email_tratado)
                    pag_reg.cell(row=row, column=8, value=os.getlogin())
                    pag_reg.cell(row=row, column=9, value=data_envio)
                    pag_reg.cell(row=row, column=10, value=os.getlogin())
                    pag_reg.cell(row=row, column=11, value=data_envio)
                    pag_reg.cell(row=row, column=12, value=data_source)
                    break
            
            wb.save(registro_path)
            
        except Exception as e:
            print(f"Erro ao salvar registro: {e}")
        
        if request.headers.get('X-Requested-With') == 'XMLHttpRequest' or request.is_json:
            return jsonify({'success': True, 'message': 'Processado com sucesso!'})
        
        return render_template("upload.html")
    
    return render_template("upload.html")

@app.route('/teste', methods=['GET', 'POST'])
def teste():
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    
    if request.method == 'POST':
        data_envio = datetime.now().strftime('%d/%m/%Y %H:%M')
        
        caminho_arquivo = request.form.get('Arquivo', '')
        folha = request.form.get('Folha', '')
        coluna_email = request.form.get('Nome Coluna E-mail', '')
        coluna_nome_destinatario = request.form.get('Nome Coluna Destinatário', '')
        colunas_salvas = request.form.get('Colunas utilizadas', '')
        assunto_original = request.form.get('Assunto', '')
        corpo = request.form.get('Corpo', '')
        data_source = request.form.get('Origem Dados', '')
        
        colunas_filtradas = []
        if colunas_salvas:
            colunas_salvas = colunas_salvas.replace(",", " ")
            colunas = colunas_salvas.split()
            for coluna in colunas:
                if coluna.strip():
                    colunas_filtradas.append(coluna.strip())
        
        user_email = get_logged_user_email()
        
        filename = os.path.basename(caminho_arquivo)
        filename_sem_extensao = filename.replace(".xlsx", "").replace(".xls", "").replace(".csv", "")
        pasta = os.path.join(UPLOAD_FOLDER, f"{filename_sem_extensao} {folha}")
        
        arquivo_encontrado = None
        possiveis_caminhos = [
            caminho_arquivo,
            os.path.join(pasta, filename),
            os.path.join(UPLOAD_FOLDER, filename),
            os.path.join(UPLOAD_FOLDER, filename_sem_extensao, filename)
        ]
        
        for caminho in possiveis_caminhos:
            if os.path.exists(caminho):
                arquivo_encontrado = caminho
                break
        
        if not arquivo_encontrado:
            flash(f'Arquivo não encontrado: {filename}', 'error')
            return redirect(url_for('table_func'))

        try:
            if filename.endswith('.csv'):
                sheet_names = ['Sheet1']
                folha_encontrada = 'Sheet1'
            else:
                xl_file = pd.ExcelFile(arquivo_encontrado)
                sheet_names = xl_file.sheet_names
                
                if folha not in sheet_names:
                    folha_encontrada = None
                    for sheet in sheet_names:
                        if folha.lower() in sheet.lower():
                            folha_encontrada = sheet
                            break
                    
                    if not folha_encontrada:
                        flash(f'Folha "{folha}" não encontrada. Folhas disponíveis: {", ".join(sheet_names)}', 'error')
                        return redirect(url_for('table_func'))
                    else:
                        folha = folha_encontrada
        except Exception as e:
            flash(f'Erro ao ler o arquivo: {str(e)}', 'error')
            return redirect(url_for('table_func'))

        try:
            if filename.endswith('.csv'):
                df = pd.read_csv(arquivo_encontrado)
            else:
                df = pd.read_excel(arquivo_encontrado, sheet_name=folha)
        except Exception as e:
            flash(f'Erro ao ler o arquivo/folha: {str(e)}', 'error')
            return redirect(url_for('table_func'))
        
        if coluna_email not in df.columns:
            flash(f'Coluna de e-mail "{coluna_email}" não encontrada. Colunas disponíveis: {", ".join(df.columns)}', 'error')
            return redirect(url_for('table_func'))
        
        if coluna_nome_destinatario not in df.columns:
            flash(f'Coluna de nome "{coluna_nome_destinatario}" não encontrada. Colunas disponíveis: {", ".join(df.columns)}', 'error')
            return redirect(url_for('table_func'))
        
        colunas_filtradas_validas = []
        for col in colunas_filtradas:
            if col in df.columns:
                colunas_filtradas_validas.append(col)
            else:
                print(f'Aviso: Coluna "{col}" não encontrada')
        
        df_sem_duplicatas = df.drop_duplicates(subset=[coluna_email], keep='first')
        
        def enviar_email_teste(filtro_df, nome_destinatario, row_email, idx):
            try:
                # Substitui placeholder no assunto
                assunto_processado = assunto_original.replace('{destinatario}', nome_destinatario or '')
                
                # Gera tabela HTML para o dataframe filtrado
                if not filtro_df.empty and colunas_filtradas_validas:
                    html_table = df_to_outlook_table(
                        filtro_df[colunas_filtradas_validas],
                        header_bg="#004b2c",
                        header_color="#ffffff",
                        zebra_even_bg="#f3f3f3"
                    )
                else:
                    html_table = "<p>Nenhum dado encontrado</p>"
                
                # CORREÇÃO CRÍTICA: Usa filtro_df (apenas dados do destinatário) para os cálculos
                corpo_processado = corpo
                if not filtro_df.empty:
                    corpo_processado = processar_calculos(corpo_processado, filtro_df)
                else:
                    corpo_processado = re.sub(r'\{[^}]+\}', '', corpo_processado)
                
                # Substitui placeholders no corpo
                corpo_final = corpo_processado.replace('{destinatario}', nome_destinatario or '')
                corpo_final = corpo_final.replace('{dataframe}', html_table)
                corpo_final = re.sub(r'\{[^}]+\}', '', corpo_final)
                
                mail = outlook.CreateItem(0)
                mail.Subject = assunto_processado
                
                if os.getlogin()[0] == 'f':
                    mail.SentOnBehalfOfName = '1751@bnb.gov.br'
                
                mail.To = user_email
                
                email_html = render_template('placeholders.html',
                    corpo_mensagem=corpo_final,
                    tabela_dataframe=''
                )
                mail.HTMLBody = email_html
                
                # Anexa arquivo se necessário
                if data_source in ['file', 'both'] and row_email is not None:
                    nome_arquivo = row_email[coluna_nome_destinatario] if coluna_nome_destinatario in row_email else nome_destinatario
                    if isinstance(nome_arquivo, str):
                        nome_arquivo = nome_arquivo.rstrip()
                    nome_arquivo = re.sub(r'[\\/*?:"<>|]', '', str(nome_arquivo))
                    arquivo_anexo = os.path.join(pasta, f"{nome_arquivo} - {filename}")
                    
                    if not os.path.exists(pasta):
                        os.makedirs(pasta)
                    
                    if not filtro_df.empty and colunas_filtradas_validas:
                        filtro_df[colunas_filtradas_validas].to_excel(arquivo_anexo, index=False)
                    
                    if os.path.exists(arquivo_anexo):
                        mail.Attachments.Add(arquivo_anexo)
                
                mail.Send()
                return True
                
            except Exception as e:
                print(f"Erro no teste para {nome_destinatario}: {e}")
                return False
        
        envios_realizados = 0
        limite = min(5, len(df_sem_duplicatas))
        
        for idx, row_email in df_sem_duplicatas.iloc[0:limite].iterrows():
            # Filtra apenas os dados do destinatário atual
            filtro = df[df[coluna_email] == str(row_email[coluna_email])]
            if not filtro.empty and colunas_filtradas_validas:
                filtro = filtro[colunas_filtradas_validas]
            else:
                filtro = pd.DataFrame()
            
            nome = row_email[coluna_nome_destinatario] if coluna_nome_destinatario in row_email else ""
            if isinstance(nome, str):
                nome = nome.rstrip()
            else:
                nome = str(nome) if nome else ""

            if enviar_email_teste(filtro, nome, row_email, idx):
                envios_realizados += 1
        
        flash(f'Teste concluído! {envios_realizados} e-mails enviados para {user_email}', 'success')
    
    return redirect(url_for('table_func'))

@app.route('/get_columns', methods=['POST'])
def get_columns():
    file = request.files['file']
    sheet_name = request.form.get('sheet_name')
   
    try:
        wb = load_workbook(file, read_only=True)
        sheet = wb[sheet_name]
       
        columns = []
        for cell in sheet[1]:
            if cell.value:
                columns.append(cell.value)
       
        return jsonify({'columns': columns})
   
    except Exception as e:
        return jsonify({'error': str(e)}), 400
       
@app.route('/get_sheets', methods=['POST'])
def get_sheets():
    if 'file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400

    file = request.files['file']
   
    if file.filename == '':
        return jsonify({'error': 'Nenhum arquivo selecionado'}), 400

    if not allowed_file(file.filename):
        return jsonify({'error': 'Tipo de arquivo não permitido'}), 400

    try:
        wb = load_workbook(file,read_only=True)
        return jsonify({'sheets': wb.sheetnames, 'active_sheet': wb.active.title})
   
    except Exception as e:
        return jsonify({'error': f'Erro ao ler o arquivo: {str(e)}'}), 500
       
@app.route('/table', methods = ['GET', 'POST'])
def table_func():
    registro = pd.read_excel(f'{UPLOAD_FOLDER}/Registro/Registro.xlsx', index_col = None)
    dados = registro.to_dict(orient="records")
   
    usuarios = registro['Criado Por'].dropna().unique().tolist() if 'Criado Por' in registro.columns else []
   
    return render_template("table.html", dados=dados, users=usuarios, current_user=os.getlogin())


@app.route('/processar', methods=['GET', 'POST'])
def processar():
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
    
    if request.method == 'POST':
        data_envio = datetime.now().strftime('%d/%m/%Y %H:%M')
        
        caminho_arquivo = request.form.get('Arquivo', '')
        folha = request.form.get('Folha', '')
        coluna_email = request.form.get('Nome Coluna E-mail', '')
        coluna_nome_destinatario = request.form.get('Nome Coluna Destinatário', '')
        colunas_salvas = request.form.get('Colunas utilizadas', '')
        assunto_original = request.form.get('Assunto', '')
        corpo = request.form.get('Corpo', '')
        data_source = request.form.get('Origem Dados', '')
        data_criacao = request.form.get('data_criacao')
                  
        corpo_processado = corpo
                          
        registro_df = pd.read_excel(f'{UPLOAD_FOLDER}/Registro/Registro.xlsx')
        registro_df.columns = registro_df.columns.str.strip()
        
        linha_encontrada = None

        for idx, row in registro_df.iterrows():
            if (
                row['Caminho Arquivo'] == caminho_arquivo and
                row['Data Criação'] == data_criacao
            ):
                linha_encontrada = idx
                break
        
        colunas_filtradas = []
        if colunas_salvas:
            colunas_salvas = colunas_salvas.replace(",", " ")
            colunas = colunas_salvas.split()
            for coluna in colunas:
                if coluna.strip():
                    colunas_filtradas.append(coluna.strip())

        caminho_arquivo_normalizado = normalizar_caminho(caminho_arquivo)
        filename = os.path.basename(caminho_arquivo_normalizado)
        filename_sem_extensao = filename.replace(".xlsx", "").replace(".xls", "").replace(".csv", "")
        pasta = normalizar_caminho(os.path.join(UPLOAD_FOLDER, f"{filename_sem_extensao} {folha}"))
        
        # CORREÇÃO 1: Encontrar o arquivo nos possíveis caminhos
        arquivo_encontrado = None
        possiveis_caminhos = [
            caminho_arquivo_normalizado,
            os.path.join(pasta, filename),
            os.path.join(UPLOAD_FOLDER, filename),
            os.path.join(UPLOAD_FOLDER, filename_sem_extensao, filename)
        ]
        
        for caminho in possiveis_caminhos:
            if os.path.exists(caminho):
                arquivo_encontrado = caminho
                break
        
        if not arquivo_encontrado:
            flash(f'Arquivo não encontrado: {filename}', 'error')
            return redirect(url_for('table_func'))
        
        try:
            if filename.endswith('.csv'):
                sheet_names = ['Sheet1']
                folha_encontrada = 'Sheet1'
            else:
                xl_file = pd.ExcelFile(arquivo_encontrado)
                sheet_names = xl_file.sheet_names
                
                if folha not in sheet_names:
                    folha_encontrada = None
                    for sheet in sheet_names:
                        if folha.lower() in sheet.lower():
                            folha_encontrada = sheet
                            break
                    
                    if not folha_encontrada:
                        flash(f'Folha "{folha}" não encontrada. Folhas disponíveis: {", ".join(sheet_names)}', 'error')
                        return redirect(url_for('table_func'))
                    else:
                        folha = folha_encontrada
        except Exception as e:
            flash(f'Erro ao ler o arquivo: {str(e)}', 'error')
            return redirect(url_for('table_func'))
        
        try:
            if filename.endswith('.csv'):
                df = pd.read_csv(arquivo_encontrado)
            else:
                df = pd.read_excel(arquivo_encontrado, sheet_name=folha)
        except Exception as e:
            flash(f'Erro ao ler o arquivo: {str(e)}', 'error')
            return redirect(url_for('table_func'))
        
        if coluna_email not in df.columns:
            flash(f'Coluna de e-mail "{coluna_email}" não encontrada', 'error')
            return redirect(url_for('table_func'))
        
        if coluna_nome_destinatario not in df.columns:
            flash(f'Coluna de nome "{coluna_nome_destinatario}" não encontrada', 'error')
            return redirect(url_for('table_func'))
        
        colunas_filtradas_validas = [col for col in colunas_filtradas if col in df.columns]
        
        df_sem_duplicatas = df.drop_duplicates(subset=[coluna_email], keep='first')
        
        if not os.path.exists(pasta):
            os.makedirs(pasta)
        
        def enviar_email(destinatario, nome_destinatario, filtro_df, row_email, idx):
            try:
                # CORREÇÃO 2: Substituir {destinatario} no assunto
                assunto_processado = assunto_original.replace('{destinatario}', nome_destinatario or '')
                
                if not filtro_df.empty and colunas_filtradas_validas:
                    html_table = df_to_outlook_table(
                        filtro_df[colunas_filtradas_validas],
                        header_bg="#004b2c",
                        header_color="#ffffff",
                        zebra_even_bg="#f3f3f3"
                    )
                else:
                    html_table = "<p>Nenhum dado encontrado</p>"
                
                # CORREÇÃO 3: Já está correto usando filtro_df para cálculos
                corpo_final = corpo_processado
                if filtro_df is not None and not filtro_df.empty:
                    corpo_final = processar_calculos(corpo_final, filtro_df)
                    print(f"[{idx}] Cálculos processados para {nome_destinatario}")
                else:
                    corpo_final = re.sub(r'\{[^}]+\}', '', corpo_final)
                
                corpo_final = corpo_final.replace('{destinatario}', nome_destinatario or '')
                corpo_final = corpo_final.replace('{dataframe}', html_table)
                corpo_final = re.sub(r'\{[^}]+\}', '', corpo_final)
                
                mail = outlook.CreateItem(0)
                mail.Subject = assunto_processado  # USANDO ASSUNTO COM PLACEHOLDER SUBSTITUÍDO
                mail.To = destinatario
                
                email_html = render_template('placeholders.html',
                    corpo_mensagem=corpo_final,
                    tabela_dataframe=''
                )
                mail.HTMLBody = email_html
                
                if data_source in ['file', 'both'] and row_email is not None:
                    nome_arquivo = row_email[coluna_nome_destinatario] if coluna_nome_destinatario in row_email else nome_destinatario
                    if isinstance(nome_arquivo, str):
                        nome_arquivo = nome_arquivo.rstrip()
                    nome_arquivo = re.sub(r'[\\/*?:"<>|]', '', str(nome_arquivo))
                    arquivo_anexo = os.path.join(pasta, f"{nome_arquivo} - {filename}")
                    
                    if not filtro_df.empty and colunas_filtradas_validas:
                        filtro_df[colunas_filtradas_validas].to_excel(arquivo_anexo, index=False)
                    
                    if os.path.exists(arquivo_anexo):
                        mail.Attachments.Add(arquivo_anexo)
                
                mail.Send()
                return True
                
            except Exception as e:
                print(f"Erro ao enviar para {destinatario}: {e}")
                return False
        
        envios_realizados = 0
        total = len(df_sem_duplicatas)
        print(f"Enviando {total} e-mails...")
        
        for idx, row_email in df_sem_duplicatas.iterrows():
            filtro = df[df[coluna_email] == str(row_email[coluna_email])]
            if not filtro.empty and colunas_filtradas_validas:
                filtro = filtro[colunas_filtradas_validas]
            else:
                filtro = pd.DataFrame()
            
            nome = row_email[coluna_nome_destinatario] if coluna_nome_destinatario in row_email else ""
            if isinstance(nome, str):
                nome = nome.rstrip()
            else:
                nome = str(nome) if nome else ""
            
            destinatario = row_email[coluna_email] if coluna_email in row_email else ""
            print(f"Enviando para: {destinatario} - {nome}")
            
            if enviar_email(destinatario, nome, filtro, row_email, idx):
                envios_realizados += 1
                print(f"Email {idx+1}/{total} enviado com sucesso")
            else:
                print(f"Falha ao enviar para {destinatario}")
        
        if linha_encontrada is not None:
            registro_df.loc[linha_encontrada, 'Enviado em'] = data_envio
            registro_df.loc[linha_encontrada, 'Enviado Por'] = os.getlogin()
            registro_df.to_excel(f'{UPLOAD_FOLDER}/Registro/Registro.xlsx', index=False)
                              
        flash(f'Processamento concluído! {envios_realizados} de {total} e-mails enviados.', 'success')
    
    return redirect(url_for('table_func'))

@app.route('/edit/<int:row_index>', methods=['GET'])
def edit(row_index):
    try:
        caminho_registro = f'{UPLOAD_FOLDER}/Registro/Registro.xlsx'
        registro_df = pd.read_excel(caminho_registro, index_col=None)
       
        if row_index < len(registro_df):
            registro = registro_df.iloc[row_index].to_dict()
           
            current_user = os.getlogin()
            criado_por = registro.get('Criado Por', '')
           
            if criado_por != current_user:
                flash('Você só pode editar registros que você criou!', 'error')
                return redirect(url_for('table_func'))
           
            registro_map = {
                'Arquivo': registro.get('Caminho Arquivo', ''),
                'Arquivo Original': registro.get('Caminho Arquivo', ''),
                'Folha': registro.get('Folha', ''),
                'Nome coluna e-mail': registro.get('Nome Coluna E-mail', ''),
                'Nome coluna destinatario': registro.get('Nome Coluna Destinatário', ''),
                'Colunas utilizadas': registro.get('Colunas utilizadas', ''),
                'Origem Dados': registro.get('Origem Dados', ''),
                'Assunto': registro.get('Assunto', ''),
                'Corpo': registro.get('Corpo', ''),
                'Data Criação': str(registro.get('Data Criação', ''))
            }
           
            return render_template("edit.html", registro=registro_map)
        else:
            flash('Registro não encontrado!', 'error')
            return redirect(url_for('table_func'))
           
    except Exception as e:
        flash(f'Erro ao carregar edição: {str(e)}', 'error')
        return redirect(url_for('table_func'))

@app.route('/processar_edit', methods=['GET'])
def processar_edit():
    dados_envio = session.get('dados_envio', {})
    print(dados_envio)
   
    if not dados_envio:
        flash('Nenhum dado para processar!', 'error')
        return redirect(url_for('table_func'))
   
    pythoncom.CoInitialize()
    outlook = win32.Dispatch('outlook.application')
   
    try:
        caminho_arquivo = dados_envio.get('Arquivo', '')
        folha = dados_envio.get('Folha', '')
        coluna_email = dados_envio.get('Nome coluna e-mail', '')
        coluna_nome_destinatario = dados_envio.get('Nome Coluna Destinatário', '')
        colunas_salvas = dados_envio.get('Colunas utilizadas', '')
        assunto_original = dados_envio.get('Assunto', '')
        corpo = dados_envio.get('Corpo', '')
        data_source = dados_envio.get('Origem Dados', '')
        data_criacao = dados_envio.get('Data Criação', '')  # Pode vir da sessão
        
        corpo_processado = corpo
       
        colunas_filtradas = [col.strip() for col in colunas_salvas.split(',') if col.strip()]
       
        user_email = get_logged_user_email()
       
        filename = caminho_arquivo.rsplit('\\', 1)[-1]
        filename_sem_extensao = filename.replace(".xlsx", "").replace(".xls", "").replace(".csv", "")
        pasta = f"{UPLOAD_FOLDER}/{filename_sem_extensao} {folha}"
       
        # Busca o arquivo nos possíveis caminhos
        arquivo_encontrado = None
        possiveis_caminhos = [
            caminho_arquivo,
            f"{pasta}/{filename}",
            f"{UPLOAD_FOLDER}/{filename}",
            f"{UPLOAD_FOLDER}/{filename_sem_extensao}/{filename}"
        ]
        
        for caminho in possiveis_caminhos:
            if os.path.exists(caminho):
                arquivo_encontrado = caminho
                break
       
        if not arquivo_encontrado:
            flash(f'Arquivo não encontrado: {filename}', 'error')
            return redirect(url_for('table_func'))
       
        if os.path.exists(arquivo_encontrado):
            if filename.endswith('.csv'):
                df = pd.read_csv(arquivo_encontrado)
            else:
                df = pd.read_excel(arquivo_encontrado, sheet_name=folha)

            colunas_lista = colunas_filtradas.copy()
            colunas_lista.append(coluna_email)
            colunas_lista.append(coluna_nome_destinatario)
           
            base = "base"
           
            if coluna_email in df.columns:
                df_sem_duplicatas = df.ffill().dropna(axis=0).drop_duplicates(subset=[coluna_email], keep='first')
            else:
                df_sem_duplicatas = df.ffill().dropna(axis=0)
           
            base_path = arquivo_encontrado
            if not os.path.exists(base_path):
                df_sem_duplicatas.to_excel(base_path, index=False, columns=colunas_lista, sheet_name=folha) 
           
            df_base = pd.read_excel(arquivo_encontrado, sheet_name=folha)        
            df_email = pd.read_excel(base_path, sheet_name=folha)
           
            def enviar_email(destinatario, nome_destinatario, filtro_df, row_email=None):
                try:
                    # CORREÇÃO 1: Substituir placeholder no assunto
                    assunto_processado = assunto_original.replace('{destinatario}', nome_destinatario or '')
                    
                    if not filtro_df.empty:
                        html_table = df_to_outlook_table(
                            filtro_df,
                            header_bg="#004b2c",
                            header_color="#ffffff",
                            zebra_even_bg="#f3f3f3"
                        )
                    else:
                        html_table = "<p>Nenhum dado encontrado</p>"
                              
                    corpo_final = corpo_processado

                    if filtro_df is not None and not filtro_df.empty:
                        corpo_final = processar_calculos(corpo_final, filtro_df)
                    else:
                        corpo_final = re.sub(r'\{[^}]+\}', '', corpo_final)
                   
                    corpo_final = corpo_final.replace('{destinatario}', nome_destinatario or '')
                    corpo_final = corpo_final.replace('{dataframe}', html_table)
                    corpo_final = re.sub(r'\{[^}]+\}', '', corpo_final)
                    
                    mail = outlook.CreateItem(0)
                    mail.Subject = assunto_processado  # ASSUNTO COM PLACEHOLDER SUBSTITUÍDO
                    mail.To = destinatario
                   
                    email_html = render_template('placeholders.html',
                        corpo_mensagem=corpo_final,
                        tabela_dataframe=''
                    )
                    mail.HTMLBody = email_html
                   
                    if data_source in ['file', 'both'] and row_email is not None:
                        nome_arquivo = row_email[coluna_nome_destinatario] if coluna_nome_destinatario in row_email else nome_destinatario
                        if isinstance(nome_arquivo, str):
                            nome_arquivo = nome_arquivo.rstrip()
                        nome_arquivo = re.sub(r'[\\/*?:"<>|]', '', str(nome_arquivo))
                        arquivo_anexo = f"{pasta}/{nome_arquivo} - {filename}"
                       
                        if not filtro_df.empty:
                            filtro_df.to_excel(arquivo_anexo, index=False)
                       
                        if os.path.exists(arquivo_anexo):
                            mail.Attachments.Add(arquivo_anexo)
                   
                    mail.Send()
                    return True
                except Exception as e:
                    print(f"Erro ao enviar para {destinatario}: {e}")
                    return False
           
            envios_realizados = 0
            total = len(df_sem_duplicatas)
            
            for index, row_email in df_sem_duplicatas.iterrows():
                if coluna_email in df_base.columns:
                    filtro = df_base[df_base[coluna_email] == str(row_email[coluna_email])]
                    filtro = filtro[colunas_filtradas] if not filtro.empty and colunas_filtradas else pd.DataFrame()
                else:
                    filtro = pd.DataFrame()
               
                nome = row_email[coluna_nome_destinatario] if coluna_nome_destinatario in row_email else ""
                if isinstance(nome, str):
                    nome = nome.rstrip()
                else:
                    nome = str(nome) if nome else ""
               
                destinatario = row_email[coluna_email] if coluna_email in row_email else ""
                if destinatario:
                    if enviar_email(destinatario, nome, filtro, row_email):
                        envios_realizados += 1
            
            # CORREÇÃO 2: Atualizar data de envio no registro original
            data_envio = datetime.now().strftime('%d/%m/%Y %H:%M')
            registro_path = f'{UPLOAD_FOLDER}/Registro/Registro.xlsx'
            
            if os.path.exists(registro_path):
                registro_df = pd.read_excel(registro_path, index_col=None)
                registro_df.columns = registro_df.columns.str.strip()
                
                # Procura o registro pelo caminho do arquivo e data de criação
                for idx, row in registro_df.iterrows():
                    if (row.get('Caminho Arquivo', '') == caminho_arquivo and
                        row.get('Data Criação', '') == data_criacao):
                        registro_df.loc[idx, 'Enviado em'] = data_envio
                        registro_df.loc[idx, 'Enviado Por'] = os.getlogin()
                        registro_df.to_excel(registro_path, index=False)
                        print(f"Registro atualizado: Enviado em = {data_envio}")
                        break
            
            flash(f'Processamento concluído! {envios_realizados} de {total} e-mails enviados.', 'success')
        else:
            flash(f'Arquivo não encontrado: {caminho_arquivo}', 'error')
           
    except Exception as e:
        flash(f'Erro ao processar envio: {str(e)}', 'error')
   
    session.pop('dados_envio', None)
    return redirect(url_for('table_func'))

@app.route('/edit_update', methods=['POST'])
def edit_update():
    try:
        form_action = request.form.get('form_action', 'save')
       
        dados_atualizados = {
            'Arquivo': request.form.get('caminho_arquivo', ''),
            'Arquivo Original': request.form.get('caminho_arquivo_original', ''),
            'Folha': request.form.get('folha', ''),
            'Nome Coluna E-mail': request.form.get('coluna_email', ''),
            'Nome Coluna Destinatário': request.form.get('coluna_nome_destinatario', ''),
            'Colunas utilizadas': request.form.get('colunas_utilizadas', ''),
            'Origem Dados': request.form.get('data_source', ''),
            'Assunto': request.form.get('assunto', ''),
            'Corpo': request.form.get('corpo', ''),
            'Última edição por': os.getlogin(),
            'Última edição em': datetime.now().strftime('%d/%m/%Y %H:%M'),
            'Data Criação': request.form.get('data_criacao')
        }

        caminho_registro = f'{UPLOAD_FOLDER}/Registro/Registro.xlsx'
        registro_df = pd.read_excel(caminho_registro, index_col=None)
                              
        arquivo_original = request.form.get('caminho_arquivo_original')
        data_criacao = request.form.get('data_criacao')

        linha_encontrada = None
        for idx, row in registro_df.iterrows():
            if (row['Caminho Arquivo'] == arquivo_original and row['Data Criação'] == data_criacao):
                linha_encontrada = idx
                break
       
        if linha_encontrada is not None and form_action in ['save', 'save_send']:
            try:
                registro_df.loc[linha_encontrada, 'Caminho Arquivo'] = dados_atualizados['Arquivo']
                registro_df.loc[linha_encontrada, 'Folha'] = dados_atualizados['Folha']
                registro_df.loc[linha_encontrada, 'Nome Coluna E-mail'] = dados_atualizados['Nome Coluna E-mail']
                registro_df.loc[linha_encontrada, 'Nome Coluna Destinatário'] = dados_atualizados['Nome Coluna Destinatário']
                registro_df.loc[linha_encontrada, 'Colunas utilizadas'] = dados_atualizados['Colunas utilizadas']
                registro_df.loc[linha_encontrada, 'Origem Dados'] = dados_atualizados['Origem Dados']
                registro_df.loc[linha_encontrada, 'Assunto'] = dados_atualizados['Assunto']
                registro_df.loc[linha_encontrada, 'Corpo'] = dados_atualizados['Corpo']
                # NÃO alterar a Data Criação original
            except Exception as e:
                print("erro", e)
           
            registro_df.to_excel(caminho_registro, index=False)
            flash('Registro salvo com sucesso!', 'success')
       
        if form_action in ['send', 'save_send']:
            dados_envio = {
                'Arquivo': dados_atualizados['Arquivo'],
                'Folha': dados_atualizados['Folha'],
                'Nome coluna e-mail': dados_atualizados['Nome Coluna E-mail'],
                'Nome coluna destinatario': dados_atualizados['Nome Coluna Destinatário'],
                'Colunas utilizadas': dados_atualizados['Colunas utilizadas'],
                'Origem Dados': dados_atualizados['Origem Dados'],
                'Assunto': dados_atualizados['Assunto'],
                'Corpo': dados_atualizados['Corpo'],
                'Data Criação': data_criacao  # ADICIONADO: passa a data original
            }
           
            session['dados_envio'] = dados_envio
            return redirect(url_for('processar_edit'))
       
        return redirect(url_for('table_func'))
       
    except Exception as e:
        flash(f'Erro ao salvar: {str(e)}', 'error')
        return redirect(url_for('table_func'))

@app.route('/duplicate/<int:row_index>', methods=['POST'])
def duplicate(row_index):
    """Duplica um registro para um novo usuário"""
    try:
        registro_df = pd.read_excel(f'{UPLOAD_FOLDER}/Registro/Registro.xlsx', index_col=None)
       
        if row_index < len(registro_df):
            registro_original = registro_df.iloc[row_index].to_dict()
           
            novo_registro = registro_original.copy()
            novo_registro['Criado Por'] = os.getlogin()
            novo_registro['Data Criação'] = datetime.now().strftime('%d/%m/%Y %H:%M')
           
            novo_df = pd.DataFrame([novo_registro])
            registro_df_atualizado = pd.concat([registro_df, novo_df], ignore_index=True)
            registro_df_atualizado.to_excel(f'{UPLOAD_FOLDER}/Registro/Registro.xlsx', index=False)
           
            flash('Registro duplicado com sucesso! Agora você pode editá-lo.', 'success')
        else:
            flash('Registro não encontrado!', 'error')
           
    except Exception as e:
        flash(f'Erro ao duplicar: {str(e)}', 'error')
   
    return redirect(url_for('table_func'))

if __name__ == "__main__":
    app.run()


# In[ ]:




