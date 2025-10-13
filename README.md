# Tratamento para Análise de Apuração de Pesagens IOS (Sapore CR3352 ArcelorMittal)

## Propósito

O objetivo principal deste script é automatizar a consolidação, limpeza, enriquecimento e análise dos dados brutos de pesagem de balanças (IOS) fornecidos pela unidade ArcelorMittal Tubarão (CR 3352).

O programa transforma múltiplas abas de um arquivo Excel de apuração em uma única base de dados consolidada. Além da consolidação básica, ele aplica regras de negócio para definir turnos e categorias, e utiliza um modelo de Machine Learning (Clusterização) para identificar inconsistências e erros de lançamento de pesagem.

**Resultado:** Um único arquivo Excel limpo e enriquecido com colunas que auxiliam a análise gerencial e orientam a tomada de decisão.

## Raciocínio Empregado e Estrutura do Código

A lógica do script é dividida em duas partes: Funções Auxiliares e de Análise e a Função Principal de Processamento (`tratamento_apuracoes.py`).

### 1. Funções Auxiliares e de Análise / Regras de Negócio

| Função | Propósito | Raciocínio Empregado |
| :--- | :--- | :--- |
| `encontrar_arquivo_apuracao()` | Localiza o arquivo de entrada. | Utiliza `glob` para buscar o arquivo que contenha a string `"apuracao_geral_arcelormittal"` no nome, permitindo flexibilidade no nome do arquivo original, que sempre difere em data/hora (ex: 20251013_1405). |
| `obter_data()` / `gerar_intervalo_de_datas()` | Gerencia a entrada de datas para o filtro do usuário. | Valida o formato da data (`DD/MM/YYYY`) e permite que o usuário escolha filtrar por uma data específica ou por um intervalo de datas, retornando uma lista de strings de data para o filtro do Pandas. |
| `formatar_coluna_data()` | Aplica o formato de data no arquivo Excel final. | Utiliza `openpyxl` (pós-salvamento do Pandas) para garantir que a coluna 'data' seja exibida no formato `DD/MM/YYYY` dentro do Excel, facilitando a leitura e filtros manuais. |
| `definir_categoria_preparacao()` | Classifica produtos em categorias (Proteína, Salada, Arroz, etc.). | Baseia-se em um dicionário de palavras-chave (`categorias`) e busca a primeira palavra do nome do produto. Prioriza correspondências exatas (`ARROZ`, `FEIJAO`) antes de buscar nas listas de sinônimos de cada categoria. |
| `definir_turno_da_pesagem()` | Define se a pesagem pertence ao turno almoço ou jantar. | Regras baseadas no nome do restaurante (constante `ABREM_TODO_DIA_ALMOCO_E_JANTAR`), `horario` e `etapa` (Produção Inicial, Transportada, etc.) para atribuir o turno correto, especialmente para restaurantes de operação contínua, como o central. |
| `avaliar_erros_na_pesagem()` | Identifica erros de lançamento com base em regras de horário e etapa (Pré-Cluster). | Aplica regras de horário do negócio (`REGRAS_HORARIO`) para identificar: **1.** Pesagens de Sobra Limpa muito cedo. **2.** Pesagens de Produção Inicial muito tarde. **3.** Pesagens fora da janela de Cadenciamento. **4.** Pesagens de `Z AMOSTRA` em etapas erradas. |
| `analisar_clusters_de_pesagem()` | Análise de ML (K-Means) para detectar Cadenciamento pesado como Perda por Preparação. | **1.** Filtra dados por restaurantes que **não** sejam MINIS (não tem cadenciamento) e remove "Z AMOSTRA", pois são sempre Perda por Preparação. **2.** Calcula o `intervalo_minutos` entre pesagens sucessivas do mesmo produto/restaurante. **3.** Aplica K-Means com 5 *clusters* sobre as características `pesagem` e `intervalo_minutos`. **4.** Identifica o *cluster* com o intervalo médio que tipicamente representa o Cadenciamento. **5.** Marca como "ERRO" todos os registros classificados como `PERDA POR PREPARAÇÃO` que caíram no *cluster* de Cadenciamento, indicando um erro de classificação na balança. |

### 2. Função Principal: `tratamento_apuracoes()`

Esta função orquestra todo o processo de ETL (Extração, Transformação e Carga).

1.  **Extração e Filtro Inicial:**
    * Lê o arquivo de entrada (`arquivo_entrada`) usando `CalamineWorkbook` para maior velocidade de leitura de arquivos grandes.
    * Solicita a opção de filtro por data ou intervalo.
    * Inicia um *loop* pelas abas do Excel, pulando abas com "CONSOLIDADO".

2.  **Transformação por Aba:**
    * Para cada aba, converte os dados em um DataFrame Pandas.
    * Insere as colunas `restaurante` e `balanca` com base no nome da aba.
    * Formata a coluna `data` e aplica o filtro de data/intervalo.
    * Adiciona o DataFrame processado à lista `dfs`.

3.  **Consolidação e Enriquecimento:**
    * Concatena todos os DataFrames válidos em um `df_final`.
    * Aplica as funções auxiliares e regras de negócio (`definir_turno_da_pesagem()`, `definir_categoria_preparacao()`, `avaliar_erros_na_pesagem()`, `analisar_clusters_de_pesagem()`).
    * Reorganiza as colunas na ordem final.

4.  **Salvamento e Feedback:**
    * Salva o `df_final` no novo arquivo (`arquivo_saida`).
    * Inclui tratamento de erros específicos (`PermissionError` para arquivo aberto) e erros gerais.
    * Chama `formatar_coluna_data()` para formatação final no Excel.
    * Fornece estatísticas de execução e mantém o console aberto com `input()` final.