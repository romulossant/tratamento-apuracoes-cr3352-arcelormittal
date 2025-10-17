import pandas as pd
from python_calamine import CalamineWorkbook
import time
import datetime
import glob
from openpyxl import load_workbook
# Bibliotecas de ML para a análise de Clusterização
from sklearn.cluster import KMeans
from sklearn.preprocessing import StandardScaler

# Variável global dos restaurantes que tem almoço e jantar
ABREM_TODO_DIA_ALMOCO_E_JANTAR = {
    "ACIARIA SUL", "COQUERIA", "MANUTENÇÃO CENTRAL", "MINI CONTÍNUO", 
    "MINI CONVERTEDOR", "MINI LTQ", "SUNCOKE", "CENTRAL"
}

# ============================================================================== #
#                       Funções auxiliares e de análise                          #
# ============================================================================== #

def encontrar_arquivo_apuracao():
    # Busca o arquivo de apuração na pasta raiz.
    arquivos_excel = glob.glob('*.xlsx')
    for nome_arquivo in arquivos_excel:
        if "apuracao_geral_arcelormittal" in nome_arquivo.lower():
            return nome_arquivo
    return None 

def gerar_intervalo_de_datas(data_inicial, data_final):
    # Gera uma lista de strings de data entre as datas inicial e final.
    data_inicial = datetime.datetime.strptime(data_inicial, '%d/%m/%Y')
    data_final = datetime.datetime.strptime(data_final, '%d/%m/%Y')
    
    intervalo = []
    while data_inicial <= data_final:
        intervalo.append(data_inicial.strftime('%d/%m/%Y')) 
        data_inicial += datetime.timedelta(days=1)
    
    return intervalo

def obter_data():
    while True:
        try:
            data_inicial = input("\n>>> Digite a data inicial (ex: 01/01/2001): ")
            data_inicial_formatada = datetime.datetime.strptime(data_inicial, '%d/%m/%Y')
            data_inicial_intervalo = data_inicial_formatada.strftime('%d/%m/%Y')

            deseja_intervalo = str(input("\n>>> Deseja filtrar por intervalo? (S/N): ")).lower()

            if deseja_intervalo == "s":
                data_final = input("\n>>> Insira a data final (ex: 01/01/2001): ")
                data_final_formatada = datetime.datetime.strptime(data_final, '%d/%m/%Y')
                data_final_intervalo = data_final_formatada.strftime('%d/%m/%Y')

                print(f"\nFiltrando pelo intervalo de {data_inicial_intervalo} a {data_final_intervalo}!\n")
                return gerar_intervalo_de_datas(data_inicial_intervalo, data_final_intervalo)
            else:
                print(f"\nFiltrando pela data {data_inicial_intervalo}!\n")
                return [data_inicial_intervalo]
        except ValueError:
            print("Formato ou intervalo inválidos. Utilize o formato 'dia/mes/ano'.\n")

def formatar_coluna_data(caminho_arquivo, nome_coluna='data'):
    # Formata a coluna de data no arquivo Excel final usando openpyxl
    try:
        wb = load_workbook(caminho_arquivo)
        ws = wb.active
        col_index = None
        for cell in ws[1]:
            if cell.value == nome_coluna:
                col_index = cell.column
                break
        
        if not col_index:
            print(f"Erro: coluna '{nome_coluna}' não encontrada.")
            return

        for row in ws.iter_rows(min_row=2, min_col=col_index, max_col=col_index):
            cell = row[0]
            try:
                data_obj = pd.to_datetime(cell.value, format='%d/%m/%Y', errors='coerce')
                if pd.notna(data_obj):
                    cell.value = data_obj
                    cell.number_format = 'DD/MM/YYYY'
            except (ValueError, TypeError):
                continue

        wb.save(caminho_arquivo)
        print("\nFormatação da coluna 'data' aplicada no arquivo final.\n")

    except Exception as e:
        print(f"ERRO: Ocorreu um erro ao formatar a planilha: {e}")


def definir_categoria_preparacao(preparacao):
    # Define a categoria (PROTEINA, SALADA, ARROZ, etc.) com base no nome do produto.
    preparacao = str(preparacao).upper().strip()
    categorias = {
        "GUARNICAO": ["ESPAGUETADA", "CREME", "POLENTA", "FAROFA", "QUIBEBE", "CANJIQUINHA", "PENNE", "PURE", "ESPAGUETE", "MACARRAO", 
                        "VIRADO", "CUSCUZ", "PIRAO", "NHOQUE", "PALHA"],
        "SALADA": ["SAL.", "ALFACE", "BETERRABA", "BERINGELA", "BERINJELA", "LENTILHA", "PEPINO", "TOMATE", "VAGEM", "CENOURA", "ERVILHA", 
                     "CHUCHU", "ABOBORA", "BATATA", "LEGUMES", "LEGUME", "COUVE", "SOJA", "REPOLHO", "TRIGO", "JILO", "GRAO", "BROCOLIS",
                     "ABOBRINHA", "JARDINEIRA"],
        "PROTEINA": ["COZIDO", "FRANGO", "BIFE", "FILEZINHO", "KIBINHO", "LINGUICA", "OVOS", "MERLUZA", "STROGONOFF", "CARRE", "ATUM", 
                       "CARNE", "OMELETE", "SALSICHA", "SALSICHAO", "ISCAS", "TILAPIA", "BISTECA", "HAMBURGUER", "EMPADAO", "PERNIL", "PICADINHO", 
                       "QUIBE", "FRICASSE", "BOBO", "CUBOS", "FILE", "ALMONDEGA", "ALMONDEGAS", "LOMBO", "DOBRADINHA", "GOULASH", "QUICHE", "KIBE",
                       "MOQUECA", "MOUSSAKA", "COSTELA", "FEIJOADA", "CORDON", "LASANHA", "MOELA", "OVO", "SOBRECOXA", "PATINHO", "FIGADO", 
                       "PIZZA", "COSTELINHA"],
        "SOBREMESA": ["DOCE", "MACA", "MELANCIA", "MELAO", "CHAMOUR", "LARANJA", "DELICIA", "MANJAR", "PUDIM", "TORTA", "CURAU", "GOIABADA",
                       "CHOCOLATE", "FLAN", "PAVE", "CACAROLA", "GELATINA", "PERA", "BANANADA", "BANANA", "MAMAO", "COCADA", "PE", "PICOLE", 
                       "ABACAXI", "TANGERINA", "TIRAMISSU", "BARRA"]
    }
    categorias_fixas = {"ARROZ": "ARROZ", "FEIJAO": "FEIJAO", "SUCO": "SUCO", "MOLHO": "MOLHO", "Z": preparacao}

    lista_preparacao = preparacao.split()
    if not lista_preparacao: return None 
        
    prim_nome_preparacao = lista_preparacao[0]
    
    if prim_nome_preparacao in categorias_fixas:
        return categorias_fixas[prim_nome_preparacao]

    if len(lista_preparacao) >= 2 and lista_preparacao[1] == "PALHA":
        return "GUARNICAO"

    for categoria, palavras in categorias.items():
        if prim_nome_preparacao in palavras:
            return categoria

    return None

def definir_turno_da_pesagem(restaurante, horario, etapa):
    # Define o turno (ALMOCO ou JANTAR) com base no restaurante, etapa e horário.
    etapa_upper = etapa.upper().strip() 
    
    abre_todo_dia_almoco_e_jantar = restaurante in ABREM_TODO_DIA_ALMOCO_E_JANTAR
    eh_prod_inicial = "PRODUCAO INICIAL" in etapa_upper
    eh_prod_transportada = "TRANSPORTADA" in etapa_upper

    if not abre_todo_dia_almoco_e_jantar:
        return "ALMOCO"
    
    # Regra: Produção Inicial Transportada (Regra de Inversão)
    if eh_prod_inicial and eh_prod_transportada:
        if horario > "16:00:00" or horario < "06:00:00":
            return "ALMOCO"
        return "JANTAR"

    # Regra: Produção Inicial Geral (Não Transportada)
    if eh_prod_inicial:
        if horario > "03:00:00" and horario < "13:30:00":
            return "ALMOCO"
        return "JANTAR"
    
    # Regra: Outras Etapas
    if horario > "06:00:00" and horario < "17:00:00":
        return "ALMOCO"
    return "JANTAR"

def avaliar_erros_na_pesagem(produto, etapa, horario, turno):
    # Avalia erros de pesagem com base em regras de horário e etapa (Pré-Cluster).
    REGRAS_HORARIO = {
        "JANTAR": {
            # Produção inicial com 15 minutos de tolerância após o início do horário de atendimento
            "PRODUCAO INICIAL": "19:15:00",
            # Cadenciamento com 15 minutos de tolerância antes do início do horário de atendimento
            "CADENCIAMENTO": ("18:45:00", "22:00:00")
        },
        "ALMOCO": {
            # Produção inicial com 15 minutos de tolerância após o início do horário de atendimento
            "PRODUCAO INICIAL": "11:00:00",
            # Cadenciamento com 15 minutos de tolerância antes do início do horário deatendimento
            "CADENCIAMENTO": ("10:30:00", "14:30:00")
        }
    }
    
    # Z AMOSTRA
    if produto == "Z AMOSTRA" and etapa != "PERDA POR PREPARACAO":
        return "ERRO NA PESAGEM DE AMOSTRA"
    
    if turno in REGRAS_HORARIO:
        regras_turno = REGRAS_HORARIO[turno]

        # Se a perda por preparação é lançado depois o fim do horário de atendimento
        if "PERDA POR PREPARACAO" in etapa and produto != "Z AMOSTRA":
            limite_horario = regras_turno.get("CADENCIAMENTO")
            if limite_horario and horario > limite_horario[1]:
                return "SOBRA LIMPA PESADA COMO PERDA POR PREP."
            # A checagem de Cadenciamento vs perda por poreparação é feita na função de Clusterização.
        
        # Se a sobra limpa está sendo pesada antes ou depois do início/fim do horário de atendimento
        if "SOBRA LIMPA" in etapa:
            if turno == "ALMOCO":
                if horario < "10:45:00" or horario > "17:00:00":
                    return "ERRO NA PESAGEM DE SOBRA LIMPA"
            elif turno == "JANTAR":
                if horario < "19:00:00" and horario > "06:00:00":
                    return "ERRO NA PESAGEM DE SOBRA LIMPA"
        
        # Se a produção inicial está sendo pesada apenas antes do início do horário de atendimento
        elif "PRODUCAO INICIAL" in etapa:
            if etapa == "PRODUCAO INICIAL TRANSPORTADA": return None
            limite_horario = regras_turno.get("PRODUCAO INICIAL")
            if horario > limite_horario:
                return "ERRO NA PESAGEM DE PROD. INICIAL"

        # Se o cadenciamento está sendo pesado só dentro do intervalo de atendimento
        elif "CADENCIAMENTO" in etapa:
            limite_horario = regras_turno.get("CADENCIAMENTO")
            if (horario < limite_horario[0] or horario > limite_horario[1]):
                return "ERRO NA PESAGEM DE CADENCIAMENTO"
                
    return None

def analisar_clusters_de_pesagem(df):
    # Aplica clusterização K-Means para identificar e sinalizar registros de perda por preparação que se comportam como cadenciamento.
    REGRAS_HORARIO_CLUSTERS = {
    "JANTAR": {
        "CADENCIAMENTO": ("19:00:00", "22:00:00")
    },
    "ALMOCO": {
        "CADENCIAMENTO": ("10:45:00", "14:30:00")
    }
    }
    
    # Pré-Processamento e Filtragem
    df.sort_values(by=['restaurante', 'horario'], inplace=True)
    df['horario'] = df['horario'].astype(str)
    
    # Excluir Z AMOSTRA da análise, pois sempre vai ser perda por preparação
    df_filtrado = df[df['produto'] != "Z AMOSTRA"].copy() 

    df_filtrado = df_filtrado[
        ~df_filtrado['restaurante'].str.contains('MINI', case=False, na=False)
    ].copy()

    # Coluna auxiliar para minutos desde a meia-noite
    def to_minutes(h):
        if pd.isna(h) or not isinstance(h, str) or len(h.split(':')) < 2: return None
        try:
            H, M, S = map(int, h.split(':'))
            return H * 60 + M
        except ValueError:
            return None
        
    df_filtrado['minutos_desde_meia_noite'] = df_filtrado['horario'].apply(to_minutes)
    
    # Filtrar por horário válido de cadenciamento (sempre dentro dos horários de atendimento)
    def eh_horario_cadenciamento(row):
        turno = row['turno']
        horario = row['horario']
        if turno not in REGRAS_HORARIO_CLUSTERS: return False
        limites = REGRAS_HORARIO_CLUSTERS[turno]["CADENCIAMENTO"]
        return limites[0] <= horario <= limites[1]

    df_filtrado['eh_horario_valido'] = df_filtrado.apply(eh_horario_cadenciamento, axis=1)
    df_filtrado = df_filtrado[df_filtrado['eh_horario_valido'] == True]
    
    if len(df_filtrado) < 10:
        print("AVISO: Dados insuficientes após o filtro de horário para Clusterização.")
        return df

    # Calcular o intervalo de tempo
    df_agrupado = df_filtrado.groupby(['restaurante', 'produto'])
    df_filtrado['intervalo_minutos'] = df_agrupado['minutos_desde_meia_noite'].diff()
    
    # Remoção de NaNs (primeira pesagem do grupo)
    df_limpo = df_filtrado.dropna(subset=['pesagem', 'intervalo_minutos']).copy() 

    # Treinamento e Clusterização (K-Means)
    CARACTERISTICAS = ['pesagem', 'intervalo_minutos']
    X = df_limpo[CARACTERISTICAS]

    escala = StandardScaler()
    X_escala = escala.fit_transform(X)
    
    K = 5 
    kmeans = KMeans(n_clusters=K, random_state=42, n_init=10)
    df_limpo['cluster_id'] = kmeans.fit_predict(X_escala)
    
    # Identificação do cluster de cadenciamento (menor intervalo médio)
    cluster_analise = df_limpo.groupby('cluster_id')['intervalo_minutos'].mean()
    cluster_cadenciamento_id = cluster_analise.idxmin()

    # 4. Aplicação da regra de desvio
    eh_ppp_suspeito_cad = (
        (df_limpo['etapa'].str.upper().str.contains("PERDA POR PREPARACAO")) &
        (df_limpo['cluster_id'] == cluster_cadenciamento_id)
    )
    
    df_limpo.loc[eh_ppp_suspeito_cad, 'erro_cluster'] = "CADENCIAMENTO PESADO COMO PERDA POR PREP."
    
    # Merge de volta ao DF original
    df = df.merge(df_limpo[['cluster_id', 'erro_cluster']], 
                  left_index=True, right_index=True, how='left')

    # Junta o novo erro com a coluna 'erro' existente
    df['erro'] = df.apply(
        lambda row: row['erro_cluster'] if pd.notna(row['erro_cluster']) else row['erro'], 
        axis=1
    )
    
    # Limpeza final das colunas auxiliares
    df.drop(columns=['minutos_desde_meia_noite', 'intervalo_minutos', 'cluster_id', 'erro_cluster', 'eh_horario_valido'], 
            inplace=True, errors='ignore')
    
    return df

# ============================================================================== #
#                               Função principal                                 #
# ============================================================================== #

arquivo_entrada = encontrar_arquivo_apuracao()

if arquivo_entrada:
    arquivo_saida = (f"apuracao_consolidada_{arquivo_entrada[37:].replace('.xlsx','')}.xlsx")
else:
    arquivo_saida = "apuracao_consolidada_ERRO.xlsx"

def tratar_planilha_apuracao():
    if not arquivo_entrada:
        print("ERRO: Nenhum arquivo .xlsx com 'apuracao_geral_arcelormittal' foi encontrado na pasta raiz.")
        input()
        return

    inicio = time.time()
    dfs = []
    
    try:
        wb = CalamineWorkbook.from_path(arquivo_entrada)
        print("=========================================================================================")
        print("#\tTRATAMENTO APURAÇÃO DE PESAGENS BALANÇAS IOS - SAPORE ARCELORMITTAL TUBARÃO\t#")
        print(f"#\t\tEm caso de dúvidas ou sugestões, romulo.santana@sapore.com.br\t\t#")     
        print("=========================================================================================\n")
   
        while True:
            opcao_usuario = input(">>> Deseja filtrar por data? (S/N): ").lower()
            if opcao_usuario == 's':
                data_para_filtro = obter_data()
                break
            elif opcao_usuario == 'n':
                print("\n\nProcessando todas as datas. Aguarde...\n\n")
                data_para_filtro = None
                break
            else:
                print("Opção inválida. Tente novamente.")
                continue

        nomes_abas_disponiveis = wb.sheet_names

        # Carregamento e inserção de colunas dependentes da aba, como por exemplo nome do restaurante e nome da balança.
        for aba_nome in nomes_abas_disponiveis:
            if "CONSOLIDADO" in aba_nome.upper(): continue

            print(f"Tratando aba {aba_nome}.")
            nome_planilha = wb.get_sheet_by_name(aba_nome)
            dados_wb = nome_planilha.to_python()

            if not dados_wb or len(dados_wb) < 2:
                print(f"Aba '{aba_nome}' está vazia ou contém apenas cabeçalho.")
                continue

            df = pd.DataFrame(dados_wb[1:], columns=dados_wb[0])
            colunas_a_manter = ["data", "horario", "etapa", "produto", "panela", "pesagem", "servico"]
            df = df[[col for col in colunas_a_manter if col in df.columns]]

            # Inserção de restaurante e balança
            nome_balanca = aba_nome.replace("3352 - ", "")
            nome_restaurante = aba_nome.replace("3352 - ", "").replace(" RECEB", "").replace(" HIB", "")
            if nome_restaurante in ["CENTRAL SOBRA LIMPA", "CENTRAL ACOUGUE", "CENTRAL CONFEITARIA", "CENTRAL SALADA", "CENTRAL ESTOQUE"]:
                nome_restaurante = "CENTRAL"
            
            df.insert(loc=1, column="restaurante", value=nome_restaurante)
            df.insert(loc=2, column="balanca", value=nome_balanca)

            # Formatação da coluna de data na planilha
            if 'data' in df.columns and not df['data'].empty:
                try:
                    df['data'] = pd.to_datetime(df['data'], errors='coerce')
                    df['data'] = df['data'].dt.strftime('%d/%m/%Y')
                    df.dropna(subset=['data'], inplace=True)
                except Exception as e:
                    print(f"Não foi possível formatar a coluna 'data' na planilha '{aba_nome}'. Erro: {e}")
            
            if opcao_usuario == 's':
                df = df[df['data'].isin(data_para_filtro)]

            if not df.empty:
                dfs.append(df)
            else:
                print(f"Aba '{aba_nome}' resultou em um DataFrame vazio após o processamento/filtragem")

        # INSERÇÃO DE COLUNAS NO DATAFRAME COM BASE EM FUNÇÕES
        dfs_validos = [df for df in dfs if not df.empty]
        if dfs_validos:
            df_final = pd.concat(dfs_validos, ignore_index=True)

            if 'horario' in df_final.columns:
                 df_final['horario'] = df_final['horario'].astype(str)
            else:
                 raise Exception("Dados sem coluna 'horario'.")
            
            # Coluna de turno
            df_final.insert(loc=3, column="turno", value=df_final.apply(
                lambda row: definir_turno_da_pesagem(restaurante=row['restaurante'], horario=row['horario'], etapa=row['etapa']), axis=1
            ))
            
            # Coluna de categoria
            if 'produto' in df_final.columns:
                df_final.insert(loc=6, column="categoria", value=df_final['produto'].apply(definir_categoria_preparacao))
            else:
                 df_final.insert(loc=6, column="categoria", value=None)

            # Identificação de erros com base em horários
            df_final['erro'] = df_final.apply(
                lambda row: avaliar_erros_na_pesagem(produto=row['produto'], etapa=row['etapa'], horario=row['horario'], turno=row['turno']), axis=1
            )
            
            # Análise por clusterização para erros de perda de preparação
            df_final = analisar_clusters_de_pesagem(df_final)
            
            # Definição da ordem final das colunas
            colunas_finais = [
            'data', 'restaurante', 'turno', 'balanca', 'horario', 
            'categoria', 'etapa', 'produto', 'panela', 'pesagem', 
            'servico', 'erro'
            ]

            df_final = df_final[[c for c in colunas_finais if c in df_final.columns]]

            # Salvamento
            df_final.to_excel(arquivo_saida, index=False)
            print(F"\nArquivo {arquivo_saida} criado. Aguarde...")
            formatar_coluna_data(arquivo_saida)

            fim = time.time()
            print(f"Total de linhas processadas: {len(df_final)}")
            print(f"Total de planilhas processadas: {len(dfs_validos)}")
            print(f"Tempo de execução: {(fim - inicio):.2f} segundos")

        else:
            print("Nenhuma aba válida para consolidação.")
    except PermissionError as e:
            print(f"\nErro ao sobreescrever o arquivo. Feche a planilha e tente novamente.")
    except Exception as e:
        print(f"\nErro durante a execução: {e}")
        
    input("\nPressione ENTER para sair.")

if __name__ == "__main__":
    tratar_planilha_apuracao()