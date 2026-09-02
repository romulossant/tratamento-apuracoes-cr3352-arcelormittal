# Tratamento para Análise de Apuração de Pesagens e Notas Fiscais IOS (Sapore CR3352 ArcelorMittal)

## Propósito

O objetivo principal deste script é automatizar a extração, consolidação, limpeza e enriquecimento dos dados operacionais da unidade ArcelorMittal Tubarão (CR 3352), contemplando tanto os registros de pesagem em balanças (IOS) quanto os dados de conferência de notas fiscais.

O programa lê a planilha bruta de apuração utilizando o leitor Calamine, processa as múltiplas abas de balanças e as abas específicas de notas fiscais, aplicando padronização de etapas, regras de turnos e categorização de preparações.

**Resultado:** Um único arquivo Excel consolidado contendo três abas organizadas (`pesagens`, `nota_item` e `nota_conferencia`), com as devidas colunas de datas tipadas e formatadas para exibição no Excel (`DD/MM/YYYY`), além da geração automática de relatórios de resumo em arquivo `.log`.

---

## Estrutura do Código e Raciocínio Empregado

A lógica do script está estruturada em funções modulares de suporte e uma rotina principal de processamento (`tratamento_apuracoes.py`).

### 1. Funções Auxiliares e Regras de Negócio

| Função | Propósito | Raciocínio Empregado |
| :--- | :--- | :--- |
| `encontrar_arquivo_apuracao()` | Localiza dinamicamente a planilha de entrada. | Usa `glob` para buscar o arquivo com padrão `"apuracao_geral_arcelormittal"` no diretório raiz, contornando variações de carimbo de data/hora no nome do arquivo original. |
| `obter_data()` / `gerar_intervalo_de_datas()` | Gerencia o filtro temporal solicitado ao usuário. | Valida as entradas no formato `DD/MM/YYYY` e gera uma lista sequencial de datas caso seja solicitado o filtro por intervalo. |
| `normalizar_texto()` | Limpeza e padronização de strings. | Remove acentuações e caracteres especiais usando decomposição Unicode (NFD), padronizando strings em caixa alta para cruzamentos consistentes. |
| `extrair_df_aba()` | Leitura genérica de abas via Calamine. | Carrega os dados brutos de qualquer aba diretamente para um DataFrame do Pandas em formato tabular. |
| `formatar_coluna_data()` | Aplica formatação de data nativa no Excel via OpenPyXL. | Localiza uma coluna específica de data pelo cabeçalho (em qualquer aba indicada pelo parâmetro `nome_aba`) e força a formatação para `DD/MM/YYYY`, facilitando filtros e tabelas dinâmicas. |
| `definir_categoria_preparacao()` | Categoriza itens em Proteína, Salada, Sobremesa, etc. | Classifica o produto a partir da primeira palavra da preparação baseando-se em mapeamentos exatos e listas de palavras-chave. |
| `definir_etapa()` | Normaliza a nomenclatura das etapas operacionais. | Padroniza as etapas produtivas (Produção Inicial, Transportada, Sobra Limpa, Cadenciamento, Perdas, etc.) considerando o contexto de balança e restaurante. |
| `definir_turno_da_pesagem()` | Atribui o turno da pesagem (Almoço ou Jantar). | Aplica regras horárias e operacionais considerando se o restaurante opera em regime integral (`ABREM_TODO_DIA_ALMOCO_E_JANTAR`), particularidades de refeições transportadas e etapas específicas (como resto ingesta). |
| `criar_logger()` / `gerar_resumo_pesagens()` | Auditoria e relatórios diários de fechamento. | Escreve arquivos `.log` no diretório `resumos_apuracao/` com a sumarização de peso (kg) agrupada por restaurante e etapa para cada dia processado. |

---

### 2. Fluxo da Rotina Principal: `tratar_planilha_apuracao()`

1. **Leitura e Seleção de Parâmetros:**
   - Abre a planilha de entrada via `CalamineWorkbook`.
   - Oferece ao usuário a opção de processamento integral ou filtragem por dia único/intervalo de datas.

2. **Processamento das Abas de Pesagens:**
   - Itera sobre as abas que possuem o prefixo `"3352 - "`.
   - Isola as colunas operacionais (`data`, `horario`, `etapa`, `produto`, `panela`, `pesagem`, `servico`).
   - Deriva as colunas `restaurante` e `balanca` a partir do nome da aba.
   - Aplica os filtros de data selecionados e agrupa os dados tratados em um DataFrame unificado (`df_final`).
   - Enriquece a base com as colunas `turno` e `categoria`, além de padronizar a coluna `etapa`.

3. **Processamento e Tratamento de Notas Fiscais:**
   - **`NOTA ITEM`**: Extrai e mantém estritamente as colunas:
     - `dt_emissao`, `num_nota`, `chave_nota`, `categoria_1`, `categoria_2`, `produto_estoque`, `conferido`, `qtde_nota`, `qtde_contagem`.
   - **`NOTA CONFERENCIA`**: Extrai e mantém estritamente as colunas:
     - `dt_emissao`, `num_nota`, `chave_nota`, `status_nota`, `conferencia_final`, `qtde_itens_nota`.
   - Aplica a filtragem por data nas notas fiscais caso o usuário tenha optado pelo filtro no início da execução.

4. **Gravação e Formatação Multi-Abas:**
   - Exporta os três DataFrames em um único arquivo de saída via `pd.ExcelWriter`:
     - Aba `pesagens`
     - Aba `nota_item`
     - Aba `nota_conferencia`
   - Invoca `formatar_coluna_data()` para garantir a formatação `DD/MM/YYYY` nas seguintes colunas:
     - `data` na aba `pesagens`
     - `dt_emissao` na aba `nota_item`
     - `dt_emissao` e `conferencia_final` na aba `nota_conferencia`

5. **Geração de Logs e Finalização:**
   - Gera relatórios analíticos de pesagens por restaurante e etapa para cada dia filtrado.
   - Trata erros de concorrência (`PermissionError` caso o arquivo de saída esteja aberto).
   - Apresenta as métricas de tempo e contagem de linhas no terminal.
