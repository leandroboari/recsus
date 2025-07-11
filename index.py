import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
import locale
import ftplib
import zipfile
import io
import os
import time
import threading
import datasus_dbc
import pandas as pd
from dbfread import DBF
import subprocess
from PIL import Image, ImageTk

# Definir ano e mês a partir de variáveis
database_directory = 'data'
downloads_directory = 'downloads'
sources_directory = 'sources'
output_directory = 'results'

# Definir localidade para português
locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')

# Dicionário de meses com seus valores correspondentes
months = {
	"Janeiro": "01",
	"Fevereiro": "02",
	"Março": "03",
	"Abril": "04",
	"Maio": "05",
	"Junho": "06",
	"Julho": "07",
	"Agosto": "08",
	"Setembro": "09",
	"Outubro": "10",
	"Novembro": "11",
	"Dezembro": "12"
}

# Função para limpar os logs
def clear_logs():
	log_text.config(state=tk.NORMAL)
	log_text.delete(1.0, tk.END)
	log_text.config(state=tk.DISABLED)

# Função para adicionar logs (apenas exibição, não editável)
def add_log(message):
	log_text.config(state=tk.NORMAL)
	log_text.insert(tk.END, f"{message}\n")
	log_text.see(tk.END)
	log_text.config(state=tk.DISABLED)

# Variável global para armazenar o caminho da última planilha gerada
excel_path = None

# Função para abrir a última planilha gerada
def open_excel():
	if excel_path and os.path.exists(excel_path):
		subprocess.Popen(["start", excel_path], shell=True)
	else:
		messagebox.showwarning("Aviso", "Nenhuma planilha disponível para abrir.")

# Função chamada ao selecionar mês, ano, CNES e tipo de sistema
def confirm():
	clear_logs()

	month = combo_month.get()
	month_value = months.get(month)
	year_value = combo_year.get()
	cnes_value = entry_cnes.get()
	source_value = combo_source.get()
	
	if month_value and year_value and cnes_value and source_value:
		add_log(f"\n\nData selecionada: {month_value}/{year_value}\nCNES: {cnes_value}\nFonte: {source_value}")

		# Ocultar o botão "Abrir Planilha" antes de iniciar o processo
		btn_open_excel.grid_remove()

		# Executar o processo em uma nova thread para não travar a interface
		threading.Thread(target=process_data, args=(month_value, year_value, cnes_value, source_value)).start()
	else:
		add_log("Selecione todos os campos.")

# Função que faz a coleta de dados e processamento
def process_data(month_value, year_value, cnes_value, source_value):
	global excel_path

	# Marcar o tempo inicial
	start_time = time.time()
	first_start_time = start_time

	################################################################################
	# PARTE 1 - DEFINIÇÃO DE VARIÁVEIS
	################################################################################

	################################################################################
	# PARTE 2 - COLETA DE DADOS SIGTAP ATUAL
	################################################################################

	# Configurações do FTP
	ftp_host = 'ftp2.datasus.gov.br'
	ftp_user = 'anonymous'
	ftp_pass = ''
	remote_directory = '/public/sistemas/tup/downloads/'  # Diretório no FTP

	add_log(f"\n\nCOLETA DE DADOS - SIGTAP {month_value}/{year_value}\n\n")

	add_log("Verificando se arquivo já foi baixado...")

	# Criar o padrão para o nome do arquivo baseado no ano e mês fornecidos
	arquivo_padrao = f'TabelaUnificada_{year_value}{month_value}'

	# Verificar se diretórios existem
	if not os.path.exists(downloads_directory):
		os.makedirs(downloads_directory)

	if not os.path.exists(sources_directory):
		os.makedirs(sources_directory)

	if not os.path.exists(output_directory):
		os.makedirs(output_directory)

	# Listar os arquivos locais
	arquivos_locais = os.listdir(downloads_directory)

	# Procurar por arquivos locais que correspondem ao padrão
	arquivo_encontrado_localmente = None
	for arquivo in arquivos_locais:
		if arquivo.startswith(arquivo_padrao):
			arquivo_encontrado_localmente = arquivo
			break

	if arquivo_encontrado_localmente:
		add_log(f"Arquivo encontrado localmente: {arquivo_encontrado_localmente}")
		local_zip = os.path.join(downloads_directory, arquivo_encontrado_localmente)
	else:
		add_log("Arquivo não encontrado localmente.")
		add_log("Conectando ao servidor DATASUS SIGTAP...")

		# Conectar ao servidor FTP e listar os arquivos no diretório
		ftp = ftplib.FTP(ftp_host)
		ftp.login(user=ftp_user, passwd=ftp_pass)
		ftp.cwd(remote_directory)  # Mudar para o diretório correto

		# Listar os arquivos no FTP
		arquivos_ftp = ftp.nlst()

		# Encontrar o arquivo que começa com o ano e o mês fornecidos
		arquivo_encontrado = None
		for arquivo in arquivos_ftp:
			if arquivo.startswith(arquivo_padrao):
				arquivo_encontrado = arquivo
				break

		if arquivo_encontrado:
			add_log(f"Arquivo encontrado no FTP: {arquivo_encontrado}")
			local_zip = os.path.join(downloads_directory, arquivo_encontrado)

			add_log("Baixando dados...")

			# Baixar o arquivo ZIP e salvar localmente
			with open(local_zip, 'wb') as f:
				ftp.retrbinary(f'RETR ' + arquivo_encontrado, f.write)

			add_log("Dados baixados com sucesso.")

			# Fechar a conexão FTP
			ftp.quit()
		else:
			add_log(f"Nenhum arquivo encontrado para o ano {year_value} e mês {month_value}.")
			ftp.quit()
			exit()  # Encerrar o script, pois não há mais nada a fazer

	add_log("Coletando informações...")

	# Inicializando o dicionário para armazenar os dados SIGTAP
	sigtap = {}

	# Abrir o arquivo ZIP e processar o arquivo tb_procedimento.txt
	with zipfile.ZipFile(local_zip, 'r') as zip_ref:
		if 'tb_procedimento.txt' in zip_ref.namelist():
			with zip_ref.open('tb_procedimento.txt') as file:
				for linha in io.TextIOWrapper(file, encoding='ISO-8859-1'):
					code = linha[0:10].strip()
					name = linha[10:260].strip()
					servico_hospitalar = int(linha[282:292].strip()) / 100
					servico_profissional = int(linha[303:312].strip()) / 100
					value = round(servico_hospitalar + servico_profissional, 2)
					ivr = value * 0.5
					sigtap[code] = {
						'name': name,
						'value': value,
						'ivr': ivr
					}

		if 'rl_procedimento_sia_sih.txt' in zip_ref.namelist():
			with zip_ref.open('rl_procedimento_sia_sih.txt') as file:

				for linha in io.TextIOWrapper(file, encoding='ISO-8859-1'):
					code = linha[0:10].strip()
					origem = linha[10:18].strip()

					if code in sigtap:
						sigtap[code]["origem"] = (sigtap[code].get("origem", "") + " " + origem).strip()

	add_log("Dados coletados com sucesso.")

	# Marcar o tempo final
	end_time = time.time()

	# Calcular o tempo total
	elapsed_time = end_time - start_time
	add_log(f"Processo concluído em {elapsed_time:.2f} segundos.")

	################################################################################
	# PARTE 3 - COLETA DE DADOS TUNEP
	################################################################################

	add_log("\n\nCOLETA DE DADOS - TUNEP\n\n")

	add_log("Coletando informações...")

	# Carregar o arquivo TUNEP.csv em um dicionário para consulta rápida
	tunep = {}
	local_tunep = os.path.join(sources_directory, "TUNEP.csv")

	with open(local_tunep, mode='r', encoding='ISO-8859-1') as file:
		next(file)
		for linha in file:
			partes = linha.strip().split(';')
			code = partes[0].strip()
			valor_sus = float(partes[1].replace('.', '').replace(',', '.'))
			valor_tunep = float(partes[2].replace('.', '').replace(',', '.'))

			if code not in tunep:
				tunep[code] = {} 
	
			tunep[code]["code"] = code
			tunep[code]["tunep"] = valor_tunep
			tunep[code]["sus"] = valor_sus
	
	for code, data in sigtap.items():
		if "origem" in data and data["origem"]:
			origens = data["origem"].split()
			sumTunep = 0
			sumSus = 0
			count = 0
			codTunep = ""
			for origem in origens:
				if origem in tunep:
					count += 1
					sumTunep += tunep[origem]["tunep"]
					sumSus += tunep[origem]["sus"]
					if codTunep != "":
						codTunep += " - "
					codTunep += tunep[origem]["code"]

			sigtap[code]["tunep"] = ""
			sigtap[code]["dif_tunep_sus"] = ""
			sigtap[code]["tunep_media"] = ""
			sigtap[code]["dif_tunep_sus_media"] = ""

			if(codTunep == ""):
				sigtap[code]["cod_tunep"] = ""
			else:
				sigtap[code]["cod_tunep"] = codTunep

			sigtap[code]["dif_tunep_sigtap"] = ""
			sigtap[code]["dif_tunep_sigtap_media"] = ""

			if(sumTunep > 0):
				mediaSus = sumSus / count
				mediaTunep = sumTunep / count

				valor_dif_tunep = mediaTunep - mediaSus
				if(valor_dif_tunep < 0):
					valor_dif_tunep = valor_dif_tunep * -1

				if(count > 1):
					sigtap[code]["sus_media"] = mediaSus
					sigtap[code]["tunep_media"] = mediaTunep
					sigtap[code]["dif_tunep_sus_media"] = valor_dif_tunep
					sigtap[code]["dif_tunep_sigtap_media"] = mediaTunep - sigtap[code]["value"]

					if(sigtap[code]["dif_tunep_sigtap_media"] < 0):
						sigtap[code]["dif_tunep_sigtap_media"] = ""
				else:
					sigtap[code]["sus"] = mediaSus
					sigtap[code]["tunep"] = mediaTunep
					sigtap[code]["dif_tunep_sus"] = valor_dif_tunep
					sigtap[code]["dif_tunep_sigtap"] = mediaTunep - sigtap[code]["value"]

					if(sigtap[code]["dif_tunep_sigtap"] < 0):
						sigtap[code]["dif_tunep_sigtap"] = ""

	add_log("Dados coletados com sucesso.")

	################################################################################
	# PARTE 4 - COLETA DE DADOS SIH/SIA
	################################################################################

	add_log("\n\nCOLETA DE DADOS - " + source_value + "\n\n")

	add_log("Verificando se arquivo já foi baixado...")

	arquivo_dbf = None

	# Marcar o tempo inicial
	start_time = time.time()

	# Configurações do FTP para o novo servidor e diretório
	ftp_host = 'ftp.datasus.gov.br'
	ftp_user = 'anonymous'
	ftp_pass = ''
	remote_directory = '/dissemin/publicos/' + source_value + 'SUS/200801_/Dados/'

	# Criar o padrão para o nome do arquivo baseado no ano e mês fornecidos
	arquivo_padrao = f'RDMG{year_value[2:]}{month_value}'  # Exemplo: 'RDMG2408'

	# Procurar por arquivos locais que correspondem ao padrão
	arquivo_encontrado_localmente = None
	for arquivo in arquivos_locais:
		if arquivo.startswith(arquivo_padrao):
			arquivo_encontrado_localmente = arquivo
			break

	if arquivo_encontrado_localmente:
		add_log(f"Arquivo encontrado localmente: {arquivo_encontrado_localmente}")
		local_file = os.path.join(downloads_directory, arquivo_encontrado_localmente)
		arquivo_dbf = local_file
	else:
		add_log("Arquivo não encontrado localmente. Conectando ao FTP...")

		# Conectar ao servidor FTP e listar os arquivos no diretório
		ftp = ftplib.FTP(ftp_host)
		ftp.login(user=ftp_user, passwd=ftp_pass)
		ftp.cwd(remote_directory)  # Mudar para o diretório correto

		# Listar os arquivos no FTP
		arquivos_ftp = ftp.nlst()

		# Encontrar o arquivo que começa com o ano e o mês fornecidos
		arquivo_encontrado = None
		for arquivo in arquivos_ftp:
			if arquivo.startswith(arquivo_padrao):
				arquivo_encontrado = arquivo
				break

		if arquivo_encontrado:
			add_log(f"Arquivo encontrado no FTP: {arquivo_encontrado}")
			local_file = os.path.join(downloads_directory, arquivo_encontrado)

			add_log("Baixando dados...")

			# Baixar o arquivo DBC e salvar localmente
			with open(local_file, 'wb') as f:
				ftp.retrbinary(f'RETR ' + arquivo_encontrado, f.write)

			add_log("Dados baixados com sucesso.")

			# Fechar a conexão FTP
			ftp.quit()

			add_log(f"Convertendo arquivo DBC...")

			arquivo_dbc = local_file
			arquivo_dbf = arquivo_dbc.replace('.dbc', '.dbf')

			datasus_dbc.decompress(arquivo_dbc, arquivo_dbf)
			add_log(f"DBF gerado com sucesso.")

			add_log(f"Deletando arquivo DBC...")
			os.remove(arquivo_dbc)
		else:
			add_log(f"Nenhum arquivo encontrado para o ano {year_value} e mês {month_value}.")
			ftp.quit()
			exit()  # Encerrar o script, pois não há mais nada a fazer

	# Marcar o tempo final
	end_time = time.time()

	# Calcular o tempo total
	elapsed_time = end_time - start_time
	add_log(f"Processo concluído em {elapsed_time:.2f} segundos.")

	# Verificar se o arquivo DBF foi criado
	if not arquivo_dbf or not os.path.exists(arquivo_dbf):
		add_log("Erro Fatal: Falha ao obter o arquivo DBF.")
		exit()

	################################################################################
	# PARTE 5 - GERAÇÃO DA PLANILHA
	################################################################################

	add_log("\n\nGERAÇÃO DA PLANILHA\n\n")

	add_log("Lendo arquivo DBF...")

	# Marcar o tempo inicial
	start_time = time.time()

	# Ler o arquivo DBF
	dbf_table = DBF(arquivo_dbf, load=True)

	# Converter para DataFrame pandas
	df = pd.DataFrame(iter(dbf_table))

	add_log("Aplicando filtros...")

	# Filtros: Hospital específico
	filtered = df[df['CNES'] == cnes_value][['PROC_REA', 'VAL_TOT']]

	# Ordenar por 'PROC_REA'
	filtered = filtered.sort_values(by='PROC_REA')

	# Converter 'PROC_REA' para string, mantendo formato correto
	filtered['PROC_REA'] = filtered['PROC_REA'].apply(lambda x: str(x).zfill(10).strip())

	# Agrupar os dados e somar o 'VAL_TOT' e contar a frequência
	df_agrupado = filtered.groupby('PROC_REA').agg(
		VAL_TOT=('VAL_TOT', 'sum'),
		FREQ=('PROC_REA', 'size')
	).reset_index()

	df_agrupado['DATA'] = f"{month_value}/{year_value}"
	df_agrupado['CNES'] = cnes_value
	df_agrupado['BD_SUS'] = source_value

	df_agrupado['NOME'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('name', '-'))

	df_agrupado['SIGTAP'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('value', 0))

	df_agrupado['SIGTAP_ORIGEM'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('sus', ""))
	df_agrupado['SIGTAP_ORIGEM_MEDIA'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('sus_media', ""))

	df_agrupado['TUNEP'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('tunep', ""))
	df_agrupado['COD_TUNEP'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('cod_tunep', ""))
	df_agrupado['TUNEP_MEDIA'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('tunep_media', ""))
	df_agrupado['DIF_TUNEP_SUS'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('dif_tunep_sus', ""))

	df_agrupado['TUNEP_SUS_TOTAL'] = (
		df_agrupado['FREQ'] *
		pd.to_numeric(df_agrupado['DIF_TUNEP_SUS'], errors='coerce')
	)

	df_agrupado['DIF_TUNEP_SUS_MEDIA'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('dif_tunep_sus_media', ""))

	df_agrupado['TUNEP_SUS_TOTAL_MEDIA'] = (
		df_agrupado['FREQ'] *
		pd.to_numeric(df_agrupado['DIF_TUNEP_SUS_MEDIA'], errors='coerce')
	)

	df_agrupado['DIF_TUNEP_SIGTAP'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('dif_tunep_sigtap', ""))

	df_agrupado['VALOR_TOTAL_TUNEP'] = (
		df_agrupado['FREQ'] *
		pd.to_numeric(df_agrupado['DIF_TUNEP_SIGTAP'], errors='coerce')
	)

	df_agrupado['DIF_TUNEP_SIGTAP_MEDIA'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('dif_tunep_sigtap_media', ""))

	df_agrupado['VALOR_TOTAL'] = (
		df_agrupado['FREQ'] *
		pd.to_numeric(df_agrupado['DIF_TUNEP_SIGTAP_MEDIA'], errors='coerce')
	)

	df_agrupado['VALOR_UNIT_IVR'] = (
		df_agrupado['SIGTAP'] +
		df_agrupado['SIGTAP'] / 2
	)

	df_agrupado['IVR'] = df_agrupado['PROC_REA'].apply(lambda proc: sigtap.get(proc, {}).get('ivr', ""))

	df_agrupado = df_agrupado[['CNES', 'COD_TUNEP', 'PROC_REA', 'NOME', 'DATA', 'VAL_TOT', 'FREQ', 'SIGTAP', "SIGTAP_ORIGEM", "SIGTAP_ORIGEM_MEDIA", "TUNEP", "TUNEP_MEDIA", "DIF_TUNEP_SUS", "TUNEP_SUS_TOTAL", "DIF_TUNEP_SUS_MEDIA", "TUNEP_SUS_TOTAL_MEDIA", "DIF_TUNEP_SIGTAP", "VALOR_TOTAL_TUNEP", "DIF_TUNEP_SIGTAP_MEDIA", "VALOR_TOTAL", "VALOR_UNIT_IVR", "IVR", "BD_SUS"]]

	df_agrupado['VAL_TOT'] = df_agrupado['VAL_TOT'].astype(float)
	df_agrupado['SIGTAP'] = df_agrupado['SIGTAP'].astype(float)
	df_agrupado['FREQ'] = df_agrupado['FREQ'].astype(int)

	source_header_name = f"Valor aprovado / realizado  no mês de referência do TABWIN{month_value}/{year_value}"
	sigtap_header_name = f"Valor unitário SIGTAP-SUS {month_value}/{year_value} no mês de referência"

	dif_tunep_sigtap_header_name = f"Diferença TUNEP - SIGTAP SUS {month_value}/{year_value}"
	dif_tunep_sigtap_media_header_name = f"Dif. Méd. TUNEP 2008 e SIGTAP {month_value}/{year_value}"

	valor_total_header_name = f"VALOR TOTAL - Dif. Méd TUNEP 2008 - SIGTAP {month_value}/{year_value}"

	df_agrupado.columns = ['CNES', 'Código de origem da TUNEP', 'Código do Procedimento', 'Nome do Procedimento', 'Data/Mês de Referência', source_header_name, 'Frequência / Quantidade aprovada', sigtap_header_name, "Valor unitário SIGTAP-SUS 2008", "Média do valor unitário SIGTAP-SUS 2008", "Valor unitário TUNEP 2008", "Média do valor unitário TUNEP 2008", "Diferença da TUNEP - SIGTAP-SUS 2008", "Valor Total TUNEP (Dif. TUNEP 2008 - SIGTAP-SUS 2008)", "Diferença Média TUNEP 2008 - Média SIGTAP-SUS 2008", "Valor Total TUNEP (Dif. Méd. TUNEP 2008 - Média SIGTAP-SUS 2008)", dif_tunep_sigtap_header_name, "VR TOTAL TUNEP (Diferença TUNEP 2008 - SIGTAP-SUS no mês de referência)", dif_tunep_sigtap_media_header_name, valor_total_header_name, "Valor unitário que deveria ser pago aplicando o IVR = SITAP-SUS mês de referência + 50% do SIGTAP-SUS no mês de referência", "50% do SIGTAP-SUS no mês de referência = IVR", "BD SUS"]
	
	add_log("Exportando dados para Planilha do Excel...")

	# Verificar se o diretório 'output_directory' existe, caso contrário, criar
	if not os.path.exists(output_directory):
		os.makedirs(output_directory)

	# Obter o timestamp atual
	timestamp = int(time.time())
	nome_arquivo = f'{cnes_value}-{year_value}-{month_value}-{timestamp}.xlsx'
	caminho_planilha = os.path.join(output_directory, nome_arquivo)

	# Exportar o resultado para Excel com formatação de moeda
	with pd.ExcelWriter(caminho_planilha, engine='xlsxwriter') as writer:
		
		df_agrupado.to_excel(writer, index=False, sheet_name='Resultados')
		
		# Acessar o workbook e worksheet
		workbook = writer.book
		worksheet = writer.sheets['Resultados']

		header_format = workbook.add_format({
			'bold': True,
			'align': 'center',
			'valign': 'bottom',
			'text_wrap': True,
			'bg_color': '#DCE6F1',
			'border': 1,
			'border_color': 'black'
		})
		
		# Definir os formatos
		moeda_format = workbook.add_format({'num_format': 'R$ #,##0.00', 'align': 'right', 'valign': 'vcenter', 'text_wrap': True})
		integer_format = workbook.add_format({'num_format': '0', 'align': 'center', 'valign': 'vcenter', 'text_wrap': True})
		general_format = workbook.add_format({'align': 'left', 'valign': 'vcenter', 'text_wrap': True})
		text_center_format = workbook.add_format({'align': 'center', 'valign': 'vcenter', 'text_wrap': True})

		# Configurar larguras e aplicar formatos
		worksheet.set_column('A:A', 8, text_center_format)
		worksheet.set_column('B:B', 40, text_center_format)
		worksheet.set_column('C:C', 13, text_center_format)
		worksheet.set_column('D:D', 70, general_format)
		worksheet.set_column('E:E', 20, text_center_format)
		worksheet.set_column('F:F', 20, moeda_format)
		worksheet.set_column('G:G', 20, integer_format)
		worksheet.set_column('H:H', 20, moeda_format)
		worksheet.set_column('I:I', 20, moeda_format)
		worksheet.set_column('J:J', 20, moeda_format)
		worksheet.set_column('K:K', 20, moeda_format)
		worksheet.set_column('L:L', 20, moeda_format)
		worksheet.set_column('M:M', 20, moeda_format)
		worksheet.set_column('N:N', 20, moeda_format)
		worksheet.set_column('O:O', 20, moeda_format)
		worksheet.set_column('P:P', 20, moeda_format)
		worksheet.set_column('Q:Q', 20, moeda_format)
		worksheet.set_column('R:R', 20, moeda_format)
		worksheet.set_column('S:S', 20, moeda_format)
		worksheet.set_column('T:T', 20, moeda_format)
		worksheet.set_column('U:U', 20, moeda_format)
		worksheet.set_column('V:V', 20, moeda_format)
		worksheet.set_column('W:W', 20, text_center_format)

		for col_num, value in enumerate(df_agrupado.columns.values):
			worksheet.write(0, col_num, value, header_format)

	add_log(f"Planilha exportada com sucesso em \"{caminho_planilha}\".")

	excel_path = caminho_planilha

	# Exibir o botão "Abrir Planilha" após gerar a planilha com sucesso
	btn_open_excel.grid()

	# Marcar o tempo final
	end_time = time.time()

	# Calcular o tempo dessa tarefa
	elapsed_time = end_time - start_time
	add_log(f"Processo concluído em {elapsed_time:.2f} segundos.")

	# Calcular o tempo total
	elapsed_time = end_time - first_start_time
	add_log(f"\nToda a operação foi concluída em {elapsed_time:.2f} segundos.\n")

# Criar janela principal
window = tk.Tk()
window.title("Procedimentos SUS")
icon_path = os.path.join(sources_directory, "itshare.ico")
window.iconbitmap(icon_path)
window.geometry("800x600")

# Frame para organizar o layout lado a lado
frame_selection = tk.Frame(window)
frame_selection.pack(pady=10)

# Selecionar Mês
label_month = tk.Label(frame_selection, text="Mês:")
label_month.grid(row=0, column=0, padx=5, pady=5)  # Adicionado pady
combo_month = ttk.Combobox(frame_selection, values=list(months.keys()), state="readonly")
combo_month.grid(row=0, column=1, padx=5, pady=5)  # Adicionado pady

# Selecionar Ano
label_year = tk.Label(frame_selection, text="Ano:")
label_year.grid(row=0, column=2, padx=5, pady=5)  # Adicionado pady
combo_year = ttk.Combobox(frame_selection, values=[str(year) for year in range(datetime.now().year + 1 - 30, datetime.now().year + 1)], state="readonly")
combo_year.grid(row=0, column=3, padx=5, pady=5)  # Adicionado pady

# Definir o mês e o ano atuais como padrão
current_month = datetime.now().strftime("%B").capitalize()
combo_month.set(current_month)
combo_year.set(str(datetime.now().year))

# Selecionar UF
label_uf = tk.Label(frame_selection, text="UF:")
label_uf.grid(row=1, column=0, padx=5, pady=5)
combo_uf = ttk.Combobox(frame_selection, values=["MG", "SP"], state="readonly")
combo_uf.grid(row=1, column=1, padx=5, pady=5)
combo_uf.set("MG")  # Definir SIH como padrão

# Campo para CNES
label_cnes = tk.Label(frame_selection, text="CNES:")
label_cnes.grid(row=1, column=2, padx=5, pady=5)
entry_cnes = tk.Entry(frame_selection)
entry_cnes.grid(row=1, column=3, padx=5, pady=5)
entry_cnes.insert(0, "2111659")  # Definir CNES padrão


# Selecionar source
label_source = tk.Label(frame_selection, text="Fonte:")
label_source.grid(row=1, column=4, padx=5, pady=5)
combo_source = ttk.Combobox(frame_selection, values=["SIH", "SIA"], state="readonly")
combo_source.grid(row=1, column=5, padx=5, pady=5)
combo_source.set("SIH")  # Definir SIH como padrão

# Campo para Índice de Correção
label_correction = tk.Label(frame_selection, text="Correção:")
label_correction.grid(row=2, column=0, padx=5, pady=5)
entry_correction = tk.Entry(frame_selection)
entry_correction.grid(row=2, column=1, padx=5, pady=5)
entry_correction.insert(0, "1.0")

# Frame para os botões
frame_buttons = tk.Frame(window)
frame_buttons.pack(pady=10)

# Botão para confirmar a seleção
btn_confirm = tk.Button(frame_buttons, text="Gerar", command=confirm)
btn_confirm.grid(row=0, column=0, padx=10)

# Botão para abrir a planilha gerada
btn_open_excel = tk.Button(frame_buttons, text="Abrir Planilha", command=open_excel)
btn_open_excel.grid(row=0, column=1, padx=10)
btn_open_excel.grid_remove()

# Campo de logs (somente leitura)
log_text = tk.Text(window, height=10, state=tk.DISABLED)
log_text.pack(padx=10, pady=10, fill=tk.BOTH, expand=True)

# Frame para a logomarca e informações do desenvolvedor
frame_footer = tk.Frame(window)
frame_footer.pack(pady=10)

try:
    logo_image = Image.open(sources_directory + "/itshare-logo-light.png")  # Substitua pelo caminho da sua logomarca
    logo_image = logo_image.resize((99, 23), Image.Resampling.LANCZOS)  # Redimensionar se necessário
    logo_photo = ImageTk.PhotoImage(logo_image)

    logo_label = tk.Label(frame_footer, image=logo_photo)
    logo_label.image = logo_photo
    logo_label.pack(side=tk.LEFT, padx=10)
except FileNotFoundError:
    add_log("Imagem da logomarca não encontrada.")

# Texto com informações do desenvolvedor
developer_info = tk.Label(
    frame_footer,
    text="ITShare Soluções em Tecnologia\nDesenvolvido por Leandro Boari Naves Silva (leandro.silva@itshare.com.br)",
    justify=tk.LEFT,
    font=("Arial", 8)
)
developer_info.pack(side=tk.LEFT, padx=10)

# Inserir mensagem inicial no campo de logs
add_log("Para começar, altere os atributos acima e clique no botão \"Gerar\".")

# Iniciar o loop da janela
window.mainloop()
