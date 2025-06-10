# Extração de Itens SEI

## Descrição

Este programa automatiza a extração de informações sobre itens de solicitação de bens de uso individual a partir de documentos no Sistema Eletrônico de Informações (SEI) da ANTT. O software acessa o sistema, navega pelos processos especificados, identifica documentos do tipo "Termo" e extrai dados estruturados de tabelas contidas nesses documentos.

## Funcionalidades Principais

### 1. Acesso Automatizado ao SEI

- Utiliza Selenium para acessar o portal SEI da ANTT (http://sei.antt.gov.br/)
- Implementa um sistema de login com autenticação via interface gráfica (módulo login_sei)
- Executa o processamento em uma thread separada para não bloquear a interface

### 2. Processamento de Múltiplos Processos

- Lê números de processos a partir de uma planilha Excel existente
- Processa cada processo de forma sequencial
- Implementa tratamento de erros para garantir que falhas em um processo não afetem outros

### 3. Navegação Inteligente pelos Documentos

- Filtra documentos que começam com "Termo" na árvore de documentos do processo
- Navega entre frames e elementos da interface web do SEI
- Implementa espera explícita para garantir o carregamento dos elementos antes de interagir

### 4. Extração de Dados Estruturados

- Extrai o nome do funcionário a partir do texto do documento
- Identifica tabelas de itens solicitados e extrai:
  - Material
  - Modelo
  - Tamanho/Gênero
  - Quantidade

### 5. Armazenamento Incremental de Dados

- Salva os dados extraídos em uma planilha Excel após processar cada processo
- Implementa um sistema de backup em caso de falha no salvamento principal
- Concatena novos dados com dados existentes para preservar o histórico

### 6. Logging Abrangente

- Registra informações detalhadas sobre o processamento em um arquivo de log
- Documenta erros e sucessos para facilitar o diagnóstico de problemas
- Registra o número de itens extraídos e salvos para cada processo

## Requisitos Técnicos

- Python 3.x
- Bibliotecas: pandas, selenium, logging, threading
- ChromeDriver compatível com a versão do Google Chrome instalada
- Acesso autorizado ao sistema SEI da ANTT

## Estrutura de Arquivos

- extracao_itens-sei.py: Script principal de extração
- `login_sei.py`: Módulo para autenticação no SEI
- `excel/itens_extraidos.xlsx`: Planilha de entrada/saída com processos e itens extraídos
- `chromedriver-win64/chromedriver.exe`: Driver do Chrome para automação
- `downloads/`: Diretório para downloads (caso necessário)
- `extracao_itens_sei.log`: Arquivo de log do programa

## Como Usar

1. Prepare uma planilha Excel com uma coluna "PROCESSO SEI" contendo os números dos processos a serem extraídos
2. Execute o programa (`python extracao_itens-sei.py`)
3. Insira suas credenciais de acesso quando solicitado
4. O programa processará todos os processos listados e salvará os resultados no arquivo Excel

## Recursos de Segurança

- Salvamento incremental após cada processo para evitar perda de dados
- Sistema de backup em caso de falha no salvamento do arquivo principal
- Tratamento extensivo de exceções para garantir a continuidade do processamento

## Limitações

- Depende da estrutura atual do SEI da ANTT; mudanças na interface podem requerer ajustes
- Requer acesso autorizado ao sistema SEI
- Processa um processo por vez, o que pode resultar em tempo de execução longo para muitos processos

---

Este programa é especialmente útil para equipes administrativas que precisam consolidar informações de solicitações de bens contidas em múltiplos processos no SEI, automatizando uma tarefa que seria tediosa e propensa a erros se realizada manualmente.
