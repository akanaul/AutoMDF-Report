import pandas as pd
from datetime import datetime, time
import os
import shutil
import glob
import unicodedata
import re
from colorama import init, Fore, Style

# Inicializar colorama
init(autoreset=True)

def remover_acentos(texto):
    """Remove acentos de um texto"""
    nfd = unicodedata.normalize('NFD', texto)
    return ''.join(char for char in nfd if unicodedata.category(char) != 'Mn')

def limpar_tela():
    os.system('cls' if os.name == 'nt' else 'clear')

def encontrar_arquivo_escala():
    """Encontra o primeiro arquivo Excel que comece com 'ESCALA' na pasta"""
    arquivos = glob.glob('1.ESCALA-FIM-TURNO/ESCALA*.xlsx')
    if arquivos:
        return arquivos[0]
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
                conteudo_sem_acentos = remover_acentos(conteudo.upper())
                
                # Extrair conteúdo de PAVÃO: (com ou sem acento)
                palavras_pavao = ['PAVÃO:', 'PAVAO:']
                for palavra in palavras_pavao:
                    palavra_busca = remover_acentos(palavra.upper())
                    if palavra_busca in conteudo_sem_acentos:
                        # Encontrar a posição no texto original
                        for i, match in enumerate(re.finditer(re.escape(palavra_busca), conteudo_sem_acentos)):
                            inicio_pavao = match.start() + len(palavra_busca)
                            # Procurar próxima seção
                            palavras_proxima = ['PENDÊNCIAS:', 'PENDENCIAS:']
                            proximo_indice = -1
                            for prox_palavra in palavras_proxima:
                                prox_busca = remover_acentos(prox_palavra.upper())
                                match_prox = re.search(re.escape(prox_busca), conteudo_sem_acentos[inicio_pavao:])
                                if match_prox:
                                    proximo_indice = inicio_pavao + match_prox.start()
                                    break
                            
                            if proximo_indice != -1:
                                pavao_content = conteudo[inicio_pavao:proximo_indice].strip().upper()
                            else:
                                pavao_content = conteudo[inicio_pavao:].strip().upper()
                            break
                        if pavao_content:
                            break
                
                # Extrair conteúdo de PENDÊNCIAS: (com ou sem acento)
                palavras_pendencias = ['PENDÊNCIAS:', 'PENDENCIAS:']
                for palavra in palavras_pendencias:
                    palavra_busca = remover_acentos(palavra.upper())
                    if palavra_busca in conteudo_sem_acentos:
                        # Encontrar a posição no texto original
                        for i, match in enumerate(re.finditer(re.escape(palavra_busca), conteudo_sem_acentos)):
                            inicio_pendencias = match.start() + len(palavra_busca)
                            # Procurar próxima seção
                            palavras_proxima = ['TROCA DE CAVALO:', 'TROCA DE CAVALO:']
                            proximo_indice = -1
                            for prox_palavra in palavras_proxima:
                                prox_busca = remover_acentos(prox_palavra.upper())
                                match_prox = re.search(re.escape(prox_busca), conteudo_sem_acentos[inicio_pendencias:])
                                if match_prox:
                                    proximo_indice = inicio_pendencias + match_prox.start()
                                    break
                            
                            if proximo_indice != -1:
                                pendencias_content = conteudo[inicio_pendencias:proximo_indice].strip().upper()
                            else:
                                pendencias_content = conteudo[inicio_pendencias:].strip().upper()
                            break
                        if pendencias_content:
                            break
                
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
        # For example, let's assume we want to count the number of drivers
        num_motoristas = len(df)
        
        # Analisar colunas VIAGEM e ESCALA para contar enviados
        enviados = 0
        pavao_count = 0
        checkout_count = 0
        
        # Encontrar todas as colunas que contém "VIAGEM"
        colunas_viagem = [col for col in df.columns if 'VIAGEM' in str(col).upper()]
        
        print(f"{Fore.CYAN}📋 Total de colunas encontradas: {len(df.columns)}{Style.RESET_ALL}")
        print(f"{Fore.CYAN}📋 Colunas de viagem: {colunas_viagem}{Style.RESET_ALL}")
        
        if colunas_viagem and 'ESCALA' in df.columns:
            print(f"{Fore.YELLOW}⏳ Analisando viagens...{Style.RESET_ALL}")
            
            # Debug: mostrar primeiras 10 linhas da primeira coluna VIAGEM e ESCALA
            primeira_col_viagem = colunas_viagem[0]
            print(f"\n{Fore.CYAN}📊 Primeiras 10 linhas de {primeira_col_viagem} e ESCALA:{Style.RESET_ALL}")
            for i in range(min(10, len(df))):
                val_viagem = df.iloc[i][primeira_col_viagem]
                val_escala = df.iloc[i]['ESCALA']
                print(f"  Linha {i}: VIAGEM={repr(val_viagem)} | ESCALA={repr(val_escala)}")
            print()
            
            for index, row in df.iterrows():
                # Contar "OK" na coluna VIAGEM (maiúsculo ou minúsculo)
                for col_viagem in colunas_viagem:
                    viagem = str(row[col_viagem]).strip().upper()
                    if viagem == 'OK':
                        pavao_count += 1
                
                # Verificar se alguma das colunas VIAGEM tem "V"
                tem_v = False
                for col_viagem in colunas_viagem:
                    viagem = str(row[col_viagem]).strip().upper()
                    if viagem == 'V':
                        tem_v = True
                        break
                
                if tem_v:
                    escala = row['ESCALA']
                    # Tentar extrair hora do campo ESCALA
                    try:
                        if isinstance(escala, datetime):
                            hora = escala.time()
                        elif isinstance(escala, time):
                            hora = escala
                        elif isinstance(escala, str):
                            # Tentar parsear string no formato "00:00" ou "00:00:00"
                            escala_limpa = escala.strip()
                            hora_partes = escala_limpa.split(':')
                            
                            if len(hora_partes) >= 2:
                                horas = int(hora_partes[0])
                                minutos = int(hora_partes[1])
                                hora = time(horas, minutos)
                            else:
                                print(f"{Fore.YELLOW}  ⚠ Linha {index+2}: V encontrado mas formato de hora inválido: {repr(escala)}{Style.RESET_ALL}")
                                continue
                        elif pd.isna(escala):
                            # Campo vazio/NaN - pula
                            print(f"{Fore.YELLOW}  ⚠ Linha {index+2}: V encontrado mas ESCALA vazia{Style.RESET_ALL}")
                            continue
                        else:
                            print(f"{Fore.YELLOW}  ⚠ Linha {index+2}: V encontrado mas tipo de ESCALA desconhecido: {type(escala)} = {repr(escala)}{Style.RESET_ALL}")
                            continue
                        
                        # Verificar se está entre 00:00 e 05:20
                        if hora >= datetime.strptime('00:00', '%H:%M').time() and hora <= datetime.strptime('05:20', '%H:%M').time():
                            enviados += 1
                            print(f"{Fore.GREEN}  ✓ Linha {index+2}: V com hora {hora} - CONTADO{Style.RESET_ALL}")
                        else:
                            checkout_count += 1
                            print(f"{Fore.YELLOW}  ⚠ Linha {index+2}: V com hora {hora} - FORA DO HORÁRIO (00:00-05:20) - CHECKOUT{Style.RESET_ALL}")
                    except Exception as ex:
                        # Se não conseguir processar a hora, ignora essa linha
                        print(f"{Fore.YELLOW}  ⚠ Linha {index+2}: Erro ao processar - {ex} | ESCALA={repr(escala)}{Style.RESET_ALL}")
                        continue
            
        else:
            print(f"{Fore.RED}✗ Colunas necessárias não encontradas!{Style.RESET_ALL}")
            if not colunas_viagem:
                print(f"{Fore.RED}  Nenhuma coluna VIAGEM encontrada{Style.RESET_ALL}")
            if 'ESCALA' not in df.columns:
                print(f"{Fore.RED}  Coluna ESCALA não existe{Style.RESET_ALL}")
        
        print(f"{Fore.GREEN}✓ Planilha carregada com sucesso!{Style.RESET_ALL}")
        print(f"{Fore.CYAN}📦 Enviados encontrados: {enviados}{Style.RESET_ALL}")
    except Exception as e:
        print(f"{Fore.RED}✗ Erro ao ler planilha: {e}{Style.RESET_ALL}")
        return

    # Prepare the report content
    data_atual = datetime.now().strftime('%d/%m')
    
    # Adicionar linha PAVAO apenas se existir pelo menos um registro
    pavao_line = f"PAVAO: {pavao_count}\n" if pavao_count > 0 else ""
    
    report_content = f"""REPORT OPERAÇÃO P2 {data_atual} - {responsavel}

PLANO DO DIA: {plano_do_dia}

ENVIADAS: {enviados}
{pavao_line}


AGUARDANDO MDF: {aguardando_mdf}
AGUARDANDO FATURAMENTO: {aguardando_faturamento}
AGUARDANDO CHECKOUT: {checkout_count}

PAVÃO:
{pavao_content}


PENDÊNCIAS:
{pendencias_content}



TROCA DE CAVALO:

MOVIMENTAÇÕES SÓ CAVALO:

ENTRADA DHL X ITU: 
SAÍDA ITU X DHL: 
SAÍDA ITU x SOROCABA: 
ENTRADA SOROCABA x ITU: 

MOTORISTA EM ATRASO:


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
        print(f"{Fore.YELLOW}⚠ Aviso ao copiar escala: {e}{Style.RESET_ALL}")
    
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