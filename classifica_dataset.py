import os
import re
import glob
import h5py
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
import datetime

PASTA_DADOS = "data"

ARQUIVO_RELATORIO_TXT = "relatorio_resumo_datasets.txt"
ARQUIVO_RELATORIO_XLSX = "relatorio_grupos_datasets.xlsx"

EXTENSOES_SUPORTADAS = ['.hdf', '.h5', '.csv', '.xlsx', '.xls', '.parquet']

REGRAS_CLASSIFICACAO = {
    'Powertrain & Motor (RPM, Torque, Potencia)': [
        'rpm', 'enginespeed', 'engine_speed', 'torque', 'trq', 'engine_load',
        'coolant', 'engine_temp', 'oil_pressure', 'throttle', 'throttle_pos'
    ],
    'Transmissao & Troca de Marchas': [
        'gear', 'selected_gear', 'actual_gear', 'clutch', 'clutch_pos', 'transmission'
    ],
    'Frenagem & Pedais': [
        'brake', 'brake_pressure', 'brake_pos', 'pedal', 'accpedal', 'deceleration'
    ],
    'Dinamica Longitudinal (Velocidade & Aceleracao X)': [
        'vehiclespeed', 'speed', 'v_x', 'accel_x', 'long_accel', 'wheel_speed'
    ],
    'Dinamica Lateral & Curvas (IMU / Giroscopio)': [
        'accel_y', 'accel_z', 'lat_accel', 'gyro', 'gyroscope', 'pitch', 'roll',
        'yaw', 'yaw_rate', 'steering', 'steering_angle', 'slip_angle', 'g_force', 'imu'
    ],
    'Consumo de Combustivel & Gases': [

        'fuelconsumption', 'fuel_rate', 'fuelrate', 'fuel_level', 'fuel_used',
        'l/100km', 'km/l', 'diesel', 'gasoline', 'co2', 'emissions'
    ],
    'Bateria & Sistema Eletrico (EV / Hibrido)': [
        'soc', 'battery', 'battery_temp', 'current', 'voltage', 'power',
        'energy', 'kwh', 'charging_status'
    ],
    'Navegacao, Rota & Elevacao (GPS)': [
        'latitude', 'longitude', 'lat', 'lon', 'altitude', 'elevation', 'gps',
        'gnss', 'gnss_status', 'satellites', 'heading', 'direction', 'route',
        'distance', 'track', 'tracks'
    ],
    'Gestao de Frota & Viagens (Trip Logs)': [

        'fleet', 'fleet_id', 'trip', 'trip_id', 'trip_status', 'fleet_status',
        'delivery_status', 'tour_status', 'start_time', 'end_time'
    ],
    'Perfil de Veiculo & Especificacoes Tecnicas': [
        'driver', 'driver_id', 'vehicle_id', 'vehid', 'vin', 'model', 'brand',
        'year', 'weight', 'generalized_weight[lb]', 'drive_wheels', 'engine_configuration'
    ]
}


def limpar_relatorios_antigos():
    padroes_para_remover = [
        ARQUIVO_RELATORIO_TXT,
        ARQUIVO_RELATORIO_XLSX,
        "relatorio_grupos_datasets_*.xlsx"
    ]
    for padrao in padroes_para_remover:
        for caminho in glob.glob(padrao):
            try:
                os.remove(caminho)
                print(f"[LIMPEZA] Arquivo antigo removido: '{caminho}'")
            except Exception as e:
                print(f"[AVISO] Nao foi possivel remover '{caminho}': {e}")


def extrair_colunas_hdf5(caminho):
    colunas = []
    try:
        with h5py.File(caminho, 'r') as hf:
            def callback(nome, obj):
                if isinstance(obj, h5py.Dataset):
                    colunas.append(nome.split('/')[-1])
            hf.visititems(callback)
    except Exception:
        pass
    return colunas


def extrair_colunas_excel(caminho):
    try:
        df_preview = pd.read_excel(caminho, nrows=2)
        return list(df_preview.columns)
    except Exception:
        return []


def extrair_colunas_generico(caminho):
    nome_arq = os.path.basename(caminho)
    if nome_arq.startswith("~$") or nome_arq.startswith("relatorio_"):
        return []

    ext = os.path.splitext(caminho)[1].lower()
    try:
        if ext in ['.hdf', '.h5']:
            return extrair_colunas_hdf5(caminho)
        elif ext == '.csv':
            df_preview = pd.read_csv(caminho, nrows=2, sep=None, engine='python')
            return list(df_preview.columns)
        elif ext in ['.xlsx', '.xls']:
            return extrair_colunas_excel(caminho)
        elif ext == '.parquet':
            df_preview = pd.read_parquet(caminho)
            return list(df_preview.columns)
    except Exception:
        return []
    return []



def classificar_lista_colunas(lista_colunas):
    pontuacao = {tema: 0 for tema in REGRAS_CLASSIFICACAO}
    matches_por_tema = {tema: [] for tema in REGRAS_CLASSIFICACAO}

    for col in lista_colunas:
        col_lower = str(col).lower()
        tokens = set(re.split(r'[^a-z0-9]+', col_lower)) | {col_lower}

        for tema, palavras_chave in REGRAS_CLASSIFICACAO.items():
            for kw in palavras_chave:
                kw_l = kw.lower()
                is_token_match = kw_l in tokens
                is_boundary_match = re.search(
                    rf'(?<![a-z0-9]){re.escape(kw_l)}(?![a-z0-9])', col_lower
                )
                if is_token_match or is_boundary_match:
                    pontuacao[tema] += 1
                    matches_por_tema[tema].append(col)
                    break  # uma keyword por coluna ja basta para pontuar o tema

    temas_encontrados = [tema for tema, qtd in pontuacao.items() if qtd > 0]

    if not temas_encontrados:
        categoria_principal = "Nao identificado / Outros"
    elif len(temas_encontrados) == 1:
        categoria_principal = temas_encontrados[0]
    else:
        temas_ordenados = sorted(temas_encontrados, key=lambda t: pontuacao[t], reverse=True)
        categoria_principal = f"Misto (Dominante: {temas_ordenados[0]})"

    return categoria_principal, temas_encontrados, matches_por_tema


def construir_resumo_detalhado(nome_fonte, tipo_fonte, qtd_arquivos, tamanho_mb,
                                colunas, temas, matches_por_tema):
    amostra_cols = ", ".join(colunas[:8]) if colunas else "Nenhuma coluna identificada"
    total_vars = len(colunas)

    detalhes_tematicos = []
    for t in temas:
        cols_do_tema = matches_por_tema.get(t, [])
        cols_str = ", ".join(cols_do_tema[:5])
        detalhes_tematicos.append(f"{t} (colunas: {cols_str})")

    if detalhes_tematicos:
        foco_explicativo = "; ".join(detalhes_tematicos)
    else:
        foco_explicativo = "atributos gerais de dados nao diretamente mapeados no dicionario padrao"

    if "Pasta" in tipo_fonte:
        resumo = (
            f"Conjunto estruturado em pasta reunindo {qtd_arquivos} arquivo(s) (totalizando {tamanho_mb:.2f} MB) "
            f"e {total_vars} variavel(is) unicas. Temas detectados, com as colunas que os originaram: {foco_explicativo}. "
            f"Amostra de colunas (ordem de leitura): [{amostra_cols}]."
        )
    else:
        resumo = (
            f"Arquivo de dados individual ({tamanho_mb:.2f} MB) estruturado com {total_vars} colunas/variaveis. "
            f"Temas detectados, com as colunas que os originaram: {foco_explicativo}. "
            f"Amostra de colunas (ordem de leitura): [{amostra_cols}]."
        )

    return resumo


def salvar_excel_formatado(dados_fontes, caminho_saida):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumo dos Datasets"

    headers = [
        "Dataset / Fonte",
        "Tipo de Fonte",
        "Qtd. Arquivos",
        "Tamanho (MB)",
        "Total Variaveis",
        "Categoria Dominante",
        "Subcategorias Detectadas",
        "Resumo Detalhado e Perfil dos Dados"
    ]
    ws.append(headers)

    for item in dados_fontes:
        ws.append([
            item['Fonte_Dataset'],
            item['Tipo'],
            item['Qtd_Arquivos'],
            item['Tamanho_Total_MB'],
            item['Total_Variaveis'],
            item['Classificacao_Geral'],
            item['Temas_Detectados'],
            item['Resumo_Executivo']
        ])

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Segoe UI", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Segoe UI", size=10)
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    ws.views.sheetView[0].showGridLines = True
    ws.row_dimensions[1].height = 28

    for cell in ws[1]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=len(headers)):
        ws.row_dimensions[row[0].row].height = 42
        for i, cell in enumerate(row):
            cell.font = data_font
            cell.border = thin_border
            if headers[i] in ["Qtd. Arquivos", "Total Variaveis"]:
                cell.number_format = '#,##0'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif headers[i] in ["Tamanho (MB)"]:
                cell.number_format = '#,##0.00'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif headers[i] == "Resumo Detalhado e Perfil dos Dados":
                cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    col_widths = [28, 18, 14, 15, 16, 38, 45, 80]
    for i, col_letter in enumerate([openpyxl.utils.get_column_letter(c) for c in range(1, len(headers) + 1)]):
        ws.column_dimensions[col_letter].width = col_widths[i]

    try:
        wb.save(caminho_saida)
        return caminho_saida
    except PermissionError:
        timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        novo_caminho = f"relatorio_grupos_datasets_{timestamp}.xlsx"
        wb.save(novo_caminho)
        print("\n[AVISO] A planilha original esta aberta no Excel.")
        print(f"[AVISO] O relatorio foi salvo em um novo arquivo: '{novo_caminho}'")
        return novo_caminho


def mapear_fontes_dados(pasta_base):
    if not os.path.exists(pasta_base):
        return []

    fontes = []
    for item in sorted(os.listdir(pasta_base)):
        if item.startswith('.') or item.startswith('~$') or item.startswith('relatorio_'):
            continue

        caminho_completo = os.path.join(pasta_base, item)

        if os.path.isdir(caminho_completo):
            arquivos = []
            for root, _, files in os.walk(caminho_completo):
                for file in files:
                    ext = os.path.splitext(file)[1].lower()
                    if ext in EXTENSOES_SUPORTADAS and not file.startswith('~$') and not file.startswith('relatorio_'):
                        arquivos.append(os.path.join(root, file))
            if arquivos:
                fontes.append({'nome': item, 'tipo': 'Pasta / Grupo', 'arquivos': arquivos})
        else:
            ext = os.path.splitext(item)[1].lower()
            if ext in EXTENSOES_SUPORTADAS:
                fontes.append({
                    'nome': item,
                    'tipo': f'Arquivo ({ext.upper().replace(".", "")})',
                    'arquivos': [caminho_completo]
                })

    return fontes


def analisar_pasta_dados(pasta_base):
    limpar_relatorios_antigos()
    fontes = mapear_fontes_dados(pasta_base)

    if not fontes:
        print("=" * 75)
        print(f"ATENCAO: Nenhum dataset ou pasta encontrado dentro de '{pasta_base}'.")
        print("=" * 75)
        return

    dados_fontes = []
    conteudo_txt = []

    conteudo_txt.append("=" * 85)
    conteudo_txt.append("       RELATORIO DE ANALISE DETALHADA E CONTEXTUALIZADA DE DATASETS")
    conteudo_txt.append("=" * 85 + "\n")

    for fonte in fontes:
        nome_fonte = fonte['nome']
        tipo_fonte = fonte['tipo']
        arquivos = fonte['arquivos']

        colunas_fonte = []
        for arq in arquivos:
            for c in extrair_colunas_generico(arq):
                if c not in colunas_fonte:
                    colunas_fonte.append(c)

        tamanho_total_mb = round(
            sum(os.path.getsize(arq) for arq in arquivos) / (1024 * 1024), 2
        )

        if colunas_fonte:
            cat_fonte, temas_fonte, matches_por_tema = classificar_lista_colunas(colunas_fonte)
        else:
            cat_fonte, temas_fonte, matches_por_tema = "Sem colunas / Erro", [], {}

        temas_str = " | ".join(temas_fonte) if temas_fonte else "Nenhum"

        resumo_detalhado = construir_resumo_detalhado(
            nome_fonte=nome_fonte,
            tipo_fonte=tipo_fonte,
            qtd_arquivos=len(arquivos),
            tamanho_mb=tamanho_total_mb,
            colunas=colunas_fonte,
            temas=temas_fonte,
            matches_por_tema=matches_por_tema
        )

        dados_fontes.append({
            'Fonte_Dataset': nome_fonte,
            'Tipo': tipo_fonte,
            'Qtd_Arquivos': len(arquivos),
            'Tamanho_Total_MB': tamanho_total_mb,
            'Total_Variaveis': len(colunas_fonte),
            'Classificacao_Geral': cat_fonte,
            'Temas_Detectados': temas_str,
            'Resumo_Executivo': resumo_detalhado
        })

        conteudo_txt.append(
            f"DATASET / FONTE           : {nome_fonte}\n"
            f"-------------------------------------------------------------------------------------\n"
            f"- Tipo de Fonte           : {tipo_fonte}\n"
            f"- Qtd. de Arquivos        : {len(arquivos)}\n"
            f"- Tamanho Acumulado       : {tamanho_total_mb} MB\n"
            f"- Total de Variaveis      : {len(colunas_fonte)}\n"
            f"- Categoria Dominante     : {cat_fonte}\n"
            f"- Subcategorias Mapeadas  : {temas_str}\n"
            f"- Resumo Detalhado        :\n"
            f"  {resumo_detalhado}\n"
            f"-------------------------------------------------------------------------------------\n\n"
        )

    if dados_fontes:
        with open(ARQUIVO_RELATORIO_TXT, 'w', encoding='utf-8') as f:
            f.write("\n".join(conteudo_txt))

        arquivo_salvo = salvar_excel_formatado(dados_fontes, ARQUIVO_RELATORIO_XLSX)

        print("=" * 75)
        print("ANALISE DETALHADA E CONTEXTUALIZADA CONCLUIDA COM SUCESSO!")
        print(f" Resumo em Texto : '{ARQUIVO_RELATORIO_TXT}'")
        print(f" Planilha Excel  : '{arquivo_salvo}'")
        print("=" * 75)


if __name__ == "__main__":
    analisar_pasta_dados(PASTA_DADOS)