import pandas as pd
from datetime import datetime, time
import os
import shutil
import glob
import unicodedata
import re
from colorama import init, Fore, Style

try:
    from openpyxl import load_workbook
    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

# Inicializar colorama
init(autoreset=True)

# Compilar regexes globalmente para performance
PLACA_PATTERN = re.compile(r'PLACA:\s*([A-Z0-9\-]{6,7})', re.IGNORECASE)
PLACA_QUALQUER_PATTERN = re.compile(r'(?<![A-Z0-9\-])([A-Z0-9\-]{6,7})(?![A-Z0-9\-])', re.IGNORECASE)
DEST_PATTERN = re.compile(r'([A-Z0-9\-]{6,7})')

def remover_acentos(texto):
    """Remove acentos de um texto"""
    nfd = unicodedata.normalize('NFD', texto)
    return ''.join(char for char in nfd if unicodedata.category(char) != 'Mn')

def _normalizar_com_mapa(texto):
    """
    Normaliza removendo acentos e retorna:
    - texto_normalizado
    - mapa_indices: lista onde cada índice do texto normalizado aponta para o índice correspondente no texto original
    """
    normalizado_chars = []
    mapa_indices = []
    for idx, ch in enumerate(texto):
        ch_normalizado = remover_acentos(ch)
        if not ch_normalizado:
            continue
        for _ in ch_normalizado:
            normalizado_chars.append(_.upper())
            mapa_indices.append(idx)
    return ''.join(normalizado_chars), mapa_indices

def _extrair_secao_texto(conteudo, cabecalhos_secao, cabecalhos_proxima_secao):
    """
    Extrai uma seção do texto usando busca sem acentos, preservando os índices do texto original.
    Retorna a substring original (sem o cabeçalho).
    """
    if not conteudo:
        return ""

    conteudo_normalizado, mapa_indices = _normalizar_com_mapa(conteudo)
    if not conteudo_normalizado:
        return ""

    cabecalhos_norm = [remover_acentos(cab.upper()) for cab in cabecalhos_secao]
    proximos_norm = [remover_acentos(cab.upper()) for cab in cabecalhos_proxima_secao]

    inicio_norm = -1
    cabecalho_encontrado = ""
    for cab in cabecalhos_norm:
        match = re.search(re.escape(cab), conteudo_normalizado)
        if match:
            inicio_norm = match.start() + len(cab)
            cabecalho_encontrado = cab
            break

    if inicio_norm == -1:
        return ""

    fim_norm = -1
    if proximos_norm:
        match_fim = re.search('|'.join(re.escape(p) for p in proximos_norm), conteudo_normalizado[inicio_norm:])
        if match_fim:
            fim_norm = inicio_norm + match_fim.start()

    if inicio_norm >= len(mapa_indices):
        return ""

    inicio_orig = mapa_indices[inicio_norm]
    if fim_norm != -1 and fim_norm - 1 < len(mapa_indices):
        fim_orig = mapa_indices[fim_norm - 1] + 1
        return conteudo[inicio_orig:fim_orig]
    return conteudo[inicio_orig:]

def _extrair_secao_por_linha(conteudo, cabecalhos_secao, cabecalhos_proxima_secao):
    """
    Extrai uma seção procurando cabeçalhos em linhas isoladas (sem texto após ':').
    Usa comparação sem acentos. Retorna a substring original (sem o cabeçalho).
    """
    if not conteudo:
        return ""

    cabecalhos_norm = {remover_acentos(cab.upper()).strip() for cab in cabecalhos_secao}
    proximos_norm = {remover_acentos(cab.upper()).strip() for cab in cabecalhos_proxima_secao}

    inicio_idx = None
    fim_idx = None
    offset = 0

    for linha in conteudo.splitlines(keepends=True):
        linha_norm = remover_acentos(linha.upper()).strip()
        if inicio_idx is None and linha_norm in cabecalhos_norm:
            inicio_idx = offset + len(linha)
        elif inicio_idx is not None and linha_norm in proximos_norm:
            fim_idx = offset
            break
        offset += len(linha)

    if inicio_idx is None:
        return ""
    if fim_idx is None:
        return conteudo[inicio_idx:]
    return conteudo[inicio_idx:fim_idx]


def extrair_motoristas_atraso(arquivo_excel, coluna_motorista, coluna_apresenta, coluna_escala):
    """
    Extrai motoristas em atraso que possuem ANOTAÇÕES (comentários do Excel) na coluna APRESENTA
    E onde o horário em APRESENTA é MAIOR que o horário em ESCALA
    Retorna string formatada: MOTORISTA - ESCALA: HH:MM - ANOTAÇÃO
    """
    motoristas_atraso = ""
    
    if not OPENPYXL_AVAILABLE:
        return motoristas_atraso
    
    try:
        wb = load_workbook(arquivo_excel)
        ws = wb.active
        
        # Encontrar índices das colunas
        col_motorista_idx = None
        col_apresenta_idx = None
        col_escala_idx = None
        
        for cell in ws[1]:
            if cell.value:
                if 'MOTORISTA' in str(cell.value).upper():
                    col_motorista_idx = cell.column
                if 'APRESENTA' in str(cell.value).upper():
                    col_apresenta_idx = cell.column
                if 'ESCALA' in str(cell.value).upper():
                    col_escala_idx = cell.column
        
        if not col_apresenta_idx or not col_motorista_idx or not col_escala_idx:
            wb.close()
            return motoristas_atraso
        
        # Iterar pelas linhas procurando por comentários em APRESENTA
        for row_num in range(2, ws.max_row + 1):
            cell_apresenta = ws.cell(row=row_num, column=col_apresenta_idx)
            
            # Verificar se a célula tem comentário/anotação
            if cell_apresenta.comment:
                anotacao_texto = cell_apresenta.comment.text
                
                # Obter dados da mesma linha
                cell_motorista = ws.cell(row=row_num, column=col_motorista_idx)
                cell_escala = ws.cell(row=row_num, column=col_escala_idx)
                
                motorista_val = cell_motorista.value
                escala_val = cell_escala.value
                apresenta_val = cell_apresenta.value
                
                if motorista_val and anotacao_texto and escala_val is not None and apresenta_val is not None:
                    motorista_str = str(motorista_val).strip()
                    anotacao_str = str(anotacao_texto).strip()
                    
                    # Extrair horas para comparação
                    hora_escala = None
                    hora_apresenta = None
                    
                    # Extrair hora de ESCALA
                    if isinstance(escala_val, datetime):
                        hora_escala = escala_val.time()
                    elif isinstance(escala_val, time):
                        hora_escala = escala_val
                    elif isinstance(escala_val, str):
                        try:
                            partes = escala_val.strip().split(':')
                            if len(partes) >= 2:
                                hora_escala = time(int(partes[0]), int(partes[1]))
                        except (ValueError, IndexError):
                            pass
                    
                    # Extrair hora de APRESENTA
                    if isinstance(apresenta_val, datetime):
                        hora_apresenta = apresenta_val.time()
                    elif isinstance(apresenta_val, time):
                        hora_apresenta = apresenta_val
                    elif isinstance(apresenta_val, str):
                        try:
                            partes = apresenta_val.strip().split(':')
                            if len(partes) >= 2:
                                hora_apresenta = time(int(partes[0]), int(partes[1]))
                        except (ValueError, IndexError):
                            pass
                    
                    # Verificar se APRESENTA > ESCALA (comparação de horas)
                    if hora_escala is not None and hora_apresenta is not None and hora_apresenta > hora_escala:
                        # Extrair apenas o corpo da anotação (após os :)
                        if ':' in anotacao_str:
                            anotacao_str = anotacao_str.split(':', 1)[1].strip()
                        
                        # Formatar escala
                        if isinstance(escala_val, datetime):
                            escala_str = escala_val.strftime('%H:%M')
                        elif isinstance(escala_val, time):
                            escala_str = escala_val.strftime('%H:%M')
                        else:
                            escala_str = str(escala_val).strip() if escala_val else ""
                        
                        motoristas_atraso += f"{motorista_str} - ESCALA: {escala_str} - {anotacao_str}\n"
        
        wb.close()
        
    except Exception as e:
        print(f"{Fore.YELLOW}⚠ Erro ao extrair motoristas em atraso: {e}{Style.RESET_ALL}")
    
    return motoristas_atraso


def extrair_placa_de_linha_pavao(linha):
    """
    Extrai placa de uma linha do PAVÃO.
    Aceita:
    - "PLACA: XXXXXX"
    - linha iniciando com XXXXXX (6 caracteres)
    Retorna a placa limpa (sem hífen) ou "".
    """
    if not linha:
        return ""

    # Prioriza padrão "PLACA: "
    match = PLACA_PATTERN.search(linha)
    if match:
        return match.group(1).strip().upper().replace('-', '')

    # Fallback: placa em qualquer lugar da linha (6 caracteres)
    match_qualquer = PLACA_QUALQUER_PATTERN.search(linha)
    if match_qualquer:
        return match_qualquer.group(1).strip().upper().replace('-', '')

    return ""

def extrair_placas_de_pavao(texto_pavao):
    """
    Extrai placas do texto de PAVÃO.
    Procura por "PLACA: XXXXXX" ou linha iniciando com XXXXXX.
    Retorna lista de placas limpas (sem hífen).
    """
    placas = []
    if not texto_pavao:
        return placas

    for linha in texto_pavao.splitlines():
        placa = extrair_placa_de_linha_pavao(linha)
        if placa:
            placas.append(placa)

    return placas

def processar_pavao_com_destino(pavao_content, df, colunas_comparacao, pavao_count_feito):
    """
    Remove linhas de PAVÃO que existem em DESTINO
    Extrai exatamente 6 caracteres após "PLACA: "
    Compara com o pavao_count_feito (número de OK contados)
    Retorna: (conteúdo atualizado, lista de placas removidas, aviso se houver discrepâncias)
    """
    if not pavao_content:
        return pavao_content, [], ""

    if isinstance(colunas_comparacao, str):
        colunas_comparacao = [colunas_comparacao]

    colunas_validas = [col for col in (colunas_comparacao or []) if col in df.columns]
    if not colunas_validas:
        return pavao_content, [], ""
    
    # Extrair placas do PAVÃO (somente linhas com "PLACA:")
    placas_pavao = extrair_placas_de_pavao(pavao_content)
    total_pavao_no_report = len(placas_pavao)
    if total_pavao_no_report == 0:
        return pavao_content, [], ""
    
    # Extrair placas das colunas de comparação (DESTINO, CAVALO, etc.)
    placas_destino = set()
    for col in colunas_validas:
        for dest in df[col].dropna():
            dest_str = str(dest).strip().upper()
            matches_dest = DEST_PATTERN.findall(dest_str)
            for match in matches_dest:
                placa_limpa = match.replace('-', '')
                placas_destino.add(placa_limpa)
    
    # Remover linhas do PAVÃO cuja placa está em DESTINO (com ou sem "PLACA:")
    placas_removidas = []
    linhas_pavao = pavao_content.strip().split('\n')
    linhas_atualizadas = []
    
    for linha in linhas_pavao:
        placa = extrair_placa_de_linha_pavao(linha)
        if placa and placa in placas_destino:
            placas_removidas.append(placa)
            continue
        linhas_atualizadas.append(linha)
    
    pavao_atualizado = '\n'.join(linhas_atualizadas).strip()
    
    # Avisar sobre discrepâncias
    # Comparar com o pavao_count_feito (OK contados), não com o total no report anterior
    aviso = ""
    if pavao_count_feito != len(placas_removidas):
        aviso = f"\n⚠️  [AVISO] PAVÃO: {str(pavao_count_feito).zfill(2)} PAVÕES foram feitos (OK), mas apenas {str(len(placas_removidas)).zfill(2)} foram encontrados em DESTINO. {str(pavao_count_feito - len(placas_removidas)).zfill(2)} ainda precisam ser removidas manualmente."
    
    return pavao_atualizado, placas_removidas, aviso



def limpar_tela():
    os.system('cls' if os.name == 'nt' else 'clear')

def _extrair_hora_segura(valor_escala):
    """
    Extrai hora com segurança de diferentes tipos de dados.
    Retorna tuple (hora:time ou None, sucesso:bool, mensagem:str)
    """
    try:
        if isinstance(valor_escala, datetime):
            return valor_escala.time(), True, ""
        elif isinstance(valor_escala, time):
            return valor_escala, True, ""
        elif isinstance(valor_escala, str):
            escala_limpa = valor_escala.strip()
            hora_partes = escala_limpa.split(':')
            if len(hora_partes) >= 2:
                horas = int(hora_partes[0])
                minutos = int(hora_partes[1])
                return time(horas, minutos), True, ""
            else:
                return None, False, f"formato de hora inválido"
        elif pd.isna(valor_escala):
            return None, False, "campo ESCALA vazio"
        else:
            return None, False, f"tipo desconhecido: {type(valor_escala)}"
    except Exception as ex:
        return None, False, str(ex)

def _hora_em_intervalo(hora, inicio_str='00:00', fim_str='05:20'):
    """Verifica se hora está no intervalo (inclusive)"""
    if hora is None:
        return False
    inicio = datetime.strptime(inicio_str, '%H:%M').time()
    fim = datetime.strptime(fim_str, '%H:%M').time()
    return hora >= inicio and hora <= fim

def encontrar_arquivo_escala():
    """Encontra o primeiro arquivo Excel que comece com 'ESCALA' na pasta"""
    arquivos = glob.glob('1.ESCALA-FIM-TURNO/ESCALA*.xlsx')
    if arquivos:
        return arquivos[0]
    return None

def obter_linhas_com_valores_reais(arquivo_excel, nome_coluna_frota):
    """
    Retorna índices das linhas que têm valores reais (não fórmulas) na coluna FROTA
    """
    linhas_reais = set()
    
    if not OPENPYXL_AVAILABLE:
        # Se openpyxl não está disponível, retorna todas as linhas (fallback)
        return None
    
    try:
        wb = load_workbook(arquivo_excel)
        ws = wb.active
        
        # Encontrar o índice da coluna FROTA
        coluna_frota_idx = None
        for cell in ws[1]:
            if cell.value and str(cell.value).strip().upper() == nome_coluna_frota.upper():
                coluna_frota_idx = cell.column
                break
        
        if coluna_frota_idx is None:
            print(f"{Fore.YELLOW}⚠ Coluna {nome_coluna_frota} não encontrada no header{Style.RESET_ALL}")
            wb.close()
            return None
        
        # Iterar pelas linhas e verificar se a célula tem fórmula
        for row_num in range(2, ws.max_row + 1):
            cell = ws.cell(row=row_num, column=coluna_frota_idx)
            # Se a célula não começa com "=" (não é fórmula) e tem valor
            if cell.value and not str(cell.value).startswith('='):
                linhas_reais.add(row_num)
        
        wb.close()
        return linhas_reais if linhas_reais else None
        
    except Exception as e:
        print(f"{Fore.YELLOW}⚠ Erro ao detectar fórmulas: {e}{Style.RESET_ALL}")
        return None

# Function to create the report
def create_report(plano_do_dia, responsavel, aguardando_mdf, aguardando_faturamento):
    # Read the Excel file
    try:
        print(f"{Fore.YELLOW}⏳ Lendo arquivo COLE_AQUI.txt...{Style.RESET_ALL}")
        pavao_content = ""
        pendencias_content = ""
        
        # Tentar ler o arquivo COLE_AQUI.txt
        try:
            with open('2.ULTIMO-REPORT/COLE_AQUI.txt', 'r', encoding='utf-8') as file:
                conteudo = file.read()
                
                # Extrair conteúdo de PAVÃO: (com ou sem acento) - somente cabeçalho em linha isolada
                pavao_raw = _extrair_secao_por_linha(
                    conteudo,
                    ['PAVÃO:', 'PAVAO:'],
                    ['PENDÊNCIAS:', 'PENDENCIAS:']
                )
                if pavao_raw:
                    pavao_content = pavao_raw.strip().upper()
                
                # Extrair conteúdo de PENDÊNCIAS: (com ou sem acento)
                pendencias_raw = _extrair_secao_por_linha(
                    conteudo,
                    ['PENDÊNCIAS:', 'PENDENCIAS:'],
                    ['TROCA DE CAVALO:']
                )
                if pendencias_raw:
                    pendencias_content = pendencias_raw.strip().upper()
                
                print(f"{Fore.CYAN}✓ Arquivo COLE_AQUI.txt lido com sucesso{Style.RESET_ALL}")
        except FileNotFoundError:
            print(f"{Fore.YELLOW}⚠ Arquivo COLE_AQUI.txt não encontrado, usando campos vazios{Style.RESET_ALL}")
        
        print(f"{Fore.YELLOW}⏳ Procurando planilha de escalas...{Style.RESET_ALL}")
        arquivo_escala = encontrar_arquivo_escala()
        
        if not arquivo_escala:
            print(f"{Fore.RED}✗ Nenhum arquivo ESCALA*.xlsx encontrado na pasta 1.ESCALA-FIM-TURNO{Style.RESET_ALL}")
            return
        
        print(f"{Fore.CYAN}📊 Encontrado: {os.path.basename(arquivo_escala)}{Style.RESET_ALL}")
        print(f"{Fore.YELLOW}⏳ Lendo planilha...{Style.RESET_ALL}")
        df = pd.read_excel(arquivo_escala)
        
        # Debug: mostrar nomes das colunas
        print(f"{Fore.CYAN}📋 Colunas encontradas: {list(df.columns)}{Style.RESET_ALL}")
        
        # Process the data as needed
        num_motoristas = len(df)
        
        # Analisar colunas VIAGEM e ESCALA para contar enviados
        enviados = 0
        pavao_count = 0
        checkout_count = 0
        saida_itu_dhl = 0
        troca_cavalo_content = ""
        motoristas_atraso_content = ""
        
        # Encontrar todas as colunas que contém "VIAGEM" - usar set para busca rápida
        colunas_viagem = [col for col in df.columns if 'VIAGEM' in str(col).upper()]
        colunas_viagem_set = set(colunas_viagem)
        
        # Encontrar coluna FROTA e MOTORISTA para TROCA DE CAVALO
        coluna_frota = None
        coluna_motorista = None
        for col in df.columns:
            col_upper = str(col).upper()
            if 'FROTA' in col_upper:
                coluna_frota = col
            if 'MOTORISTA' in col_upper:
                coluna_motorista = col
        
        print(f"{Fore.CYAN}📋 Colunas de viagem: {len(colunas_viagem)}{Style.RESET_ALL}")
        
        # Obter linhas com valores reais (não fórmulas) na coluna FROTA
        linhas_frota_reais = None
        if coluna_frota:
            linhas_frota_reais = obter_linhas_com_valores_reais(arquivo_escala, coluna_frota)
            if linhas_frota_reais:
                print(f"{Fore.CYAN}📦 Encontradas {len(linhas_frota_reais)} linhas com valores reais em FROTA{Style.RESET_ALL}")
        
        if colunas_viagem and 'ESCALA' in df.columns:
            print(f"{Fore.YELLOW}⏳ Analisando viagens...{Style.RESET_ALL}")
            
            # Pré-calcular tempo de referência uma única vez
            tempo_inicio = datetime.strptime('00:00', '%H:%M').time()
            tempo_fim = datetime.strptime('05:20', '%H:%M').time()
            
            # Processar cada linha uma única vez
            for index, row in df.iterrows():
                linha_excel = index + 2
                
                # Contar "OK" na coluna VIAGEM - usar any para sair rápido
                viagem_values_ok = any(str(row.get(col, '')).strip().upper() == 'OK' for col in colunas_viagem)
                if viagem_values_ok:
                    pavao_count += 1
                
                # Verificar se alguma das colunas VIAGEM tem "V" ou "SC"
                viagem_values = [str(row.get(col, '')).strip().upper() for col in colunas_viagem]
                
                # Processar ambos V e SC em uma única passada
                tem_v = 'V' in viagem_values
                tem_sc = 'SC' in viagem_values
                
                if tem_v or tem_sc:
                    escala = row['ESCALA']
                    hora, sucesso, msg = _extrair_hora_segura(escala)
                    
                    if sucesso:
                        if tem_v and _hora_em_intervalo(hora, '00:00', '05:20'):
                            enviados += 1
                        elif tem_v:
                            checkout_count += 1
                        
                        if tem_sc and _hora_em_intervalo(hora, '00:00', '05:20'):
                            saida_itu_dhl += 1
                
                # Coletar dados de TROCA DE CAVALO (valores que não são fórmulas)
                if coluna_frota and coluna_motorista and linhas_frota_reais is not None:
                    if linha_excel in linhas_frota_reais:
                        frota_val = row[coluna_frota]
                        motorista_val = row[coluna_motorista]
                        
                        frota_str = str(frota_val).strip() if pd.notna(frota_val) else ""
                        motorista_str = str(motorista_val).strip() if pd.notna(motorista_val) else ""
                        
                        frota_valida = frota_str and frota_str != 'nan' and frota_str != '-'
                        motorista_valido = motorista_str and motorista_str != 'nan' and motorista_str != '-'
                        if motorista_valido and frota_valida:
                            troca_cavalo_content += f"{motorista_str} - {frota_str}\n"
            
        else:
            print(f"{Fore.RED}✗ Colunas necessárias não encontradas!{Style.RESET_ALL}")
            if not colunas_viagem:
                print(f"{Fore.RED}  Nenhuma coluna VIAGEM encontrada{Style.RESET_ALL}")
            if 'ESCALA' not in df.columns:
                print(f"{Fore.RED}  Coluna ESCALA não existe{Style.RESET_ALL}")
        
        print(f"{Fore.GREEN}✓ Planilha carregada com sucesso!{Style.RESET_ALL}")
        print(f"{Fore.CYAN}📦 Enviados encontrados: {str(enviados).zfill(2)}{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}✗ Erro ao ler planilha: {e}{Style.RESET_ALL}")
        return

    # Prepare the report content
    data_atual = datetime.now().strftime('%d/%m')
    
    # Extrair motoristas em atraso (com anotações do Excel na coluna APRESENTA)
    motoristas_atraso_content = extrair_motoristas_atraso(arquivo_escala, coluna_motorista, 'APRESENTA', 'ESCALA')
    if motoristas_atraso_content:
        print(f"{Fore.CYAN}ℹ Motoristas em atraso encontrados com anotações:{Style.RESET_ALL}")
        for linha in motoristas_atraso_content.strip().split('\n'):
            print(f"  - {linha}")
    
    # Processar PAVÃO: remover linhas que correspondem a placas em DESTINO
    colunas_comparacao = [col for col in ['CAVALO', 'DESTINO'] if col in df.columns]
    aviso_pavao = ""
    if colunas_comparacao:
        pavao_content_processado, placas_removidas, aviso_pavao = processar_pavao_com_destino(pavao_content, df, colunas_comparacao, pavao_count)
        if placas_removidas:
            print(f"{Fore.CYAN}ℹ Removidas {str(len(placas_removidas)).zfill(2)} placa(s) do PAVÃO que foram encontradas na escala (CAVALO/DESTINO){Style.RESET_ALL}")
            for placa in placas_removidas:
                print(f"  - {placa}")
        if aviso_pavao:
            print(f"{Fore.YELLOW}{aviso_pavao}{Style.RESET_ALL}")
    else:
        pavao_content_processado = pavao_content
    
    # Adicionar linha PAVAO apenas se existir pelo menos um registro
    pavao_line = f"PAVAO: {str(pavao_count).zfill(2)}\n" if pavao_count > 0 else ""

    saida_itu_dhl_output = str(saida_itu_dhl).zfill(2) if saida_itu_dhl > 0 else ""
    
    report_content = f"""REPORT OPERAÇÃO P2 {data_atual} - {responsavel}

PLANO DO DIA: {plano_do_dia}

ENVIADAS: {str(enviados).zfill(2)}
{pavao_line}

AGUARDANDO MDF: {aguardando_mdf}
AGUARDANDO FATURAMENTO: {aguardando_faturamento}
AGUARDANDO CHECKOUT: {str(checkout_count).zfill(2)}

PAVÃO:

{pavao_content_processado}

PENDÊNCIAS:

{pendencias_content}

TROCA DE CAVALO:

{troca_cavalo_content}

MOVIMENTAÇÕES SÓ CAVALO:

ENTRADA DHL X ITU: 
SAÍDA ITU X DHL: {saida_itu_dhl_output}
SAÍDA ITU x SOROCABA: 
ENTRADA SOROCABA x ITU: 

MOTORISTA EM ATRASO:

{motoristas_atraso_content}

"""

    # Write to a new report file
    timestamp = datetime.now().strftime('%d-%m-%Y %H-%M-%S')
    data_sem_ano = datetime.now().strftime('%d-%m')
    arquivo_raiz = 'ULTIMO_RELATORIO.txt'
    arquivo_historico = f'3.HISTORICO-REPORT/REPORT {responsavel} {timestamp}.txt'
    
    # Salvar na raiz
    with open(arquivo_raiz, 'w', encoding='utf-8') as file:
        file.write(report_content)
    
    # Salvar cópia no histórico
    with open(arquivo_historico, 'w', encoding='utf-8') as file:
        file.write(report_content)
    
    # Copiar escala para histórico (sobrescreve se já existe no dia)
    nome_arquivo_escala = os.path.basename(arquivo_escala)
    nome_sem_extensao = os.path.splitext(nome_arquivo_escala)[0]
    arquivo_escala_destino = f'4.HISTORICO-ESCALA/{nome_sem_extensao} {data_sem_ano}.xlsx'
    try:
        shutil.copy2(arquivo_escala, arquivo_escala_destino)
        print(f"{Fore.GREEN}✓ Escala copiada para histórico{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.YELLOW}[AVISO] Aviso ao copiar escala: {e}{Style.RESET_ALL}")
    
    print(f"\n{Fore.GREEN}{'='*60}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}✓ RELATÓRIO CRIADO COM SUCESSO!{Style.RESET_ALL}")
    print(f"{Fore.CYAN}📄 Raiz: {arquivo_raiz}{Style.RESET_ALL}")
    print(f"{Fore.CYAN} Histórico: {arquivo_historico}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}📊 Escala: {arquivo_escala_destino}{Style.RESET_ALL}")
    print(f"{Fore.GREEN}{'='*60}{Style.RESET_ALL}\n")

# Main execution
if __name__ == '__main__':
    limpar_tela()
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{Style.BRIGHT}    GERADOR DE RELATÓRIO - OPERAÇÃO P2{Style.RESET_ALL}")
    print(f"{Fore.CYAN}{'='*60}{Style.RESET_ALL}\n")
    
    plano_do_dia = input(f"{Fore.YELLOW}📋 Qual o plano do dia? {Fore.WHITE}» {Style.RESET_ALL}").upper()
    responsavel = input(f"{Fore.YELLOW}👤 Quem é o responsável? {Fore.WHITE}» {Style.RESET_ALL}").upper()
    aguardando_mdf = input(f"{Fore.YELLOW}📦 Quantas estão aguardando MDF? {Fore.WHITE}» {Style.RESET_ALL}").upper()
    aguardando_faturamento = input(f"{Fore.YELLOW}💳 Quantas estão aguardando FATURAMENTO? {Fore.WHITE}» {Style.RESET_ALL}").upper()
    
    print(f"\n{Fore.CYAN}{'─'*60}{Style.RESET_ALL}\n")
    create_report(plano_do_dia, responsavel, aguardando_mdf, aguardando_faturamento)