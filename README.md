# 🚀 Consolidate Product Sheets

Este repositório contém um conjunto de scripts Python desenvolvidos para automatizar o processo de extração, limpeza e consolidação de dados de produtos e fornecedores a partir de planilhas Excel complexas.

O objetivo principal é transformar catálogos com múltiplos formatos e abas em arquivos CSV padronizados, prontos para importação em bancos de dados, sistemas de BI ou outras ferramentas de análise.

## ✨ Funcionalidades Principais

* **Consolidação Inteligente**: Agrega dados de múltiplas abas de um mesmo arquivo Excel em um único dataset.
* **Detecção Dinâmica de Cabeçalho**: O script de produtos localiza automaticamente a linha do cabeçalho em cada aba, oferecendo flexibilidade para planilhas com layouts variados.
* **Limpeza e Padronização Avançada**:
    * Normaliza textos, removendo acentos, convertendo para maiúsculas e eliminando caracteres especiais.
    * Converte valores numéricos no formato brasileiro (ex: `1.234,50`) para o formato decimal padrão.
    * Renomeia colunas para um padrão `snake_case`, ideal para interoperabilidade com bancos de dados.
* **Cálculo de Preço de Venda**: Aplica um fator de *markup* (margem de lucro) configurável para gerar preços de venda.
* **Modularidade**: Scripts separados para processar produtos e fornecedores, cada um com sua lógica específica.

## 📂 Estrutura do Projeto

O projeto é dividido em dois scripts principais. **Atenção:** renomeie seus arquivos `.py` para os nomes abaixo para manter a consistência com a documentação.

* `processa_produtos.py`: Script principal que varre todas as abas da planilha, consolida os dados de produtos, limpa as informações e calcula preços.
* `processa_fornecedores.py`: Script focado em extrair e limpar os dados cadastrais dos fornecedores, que geralmente se encontram em uma aba específica (a primeira, por padrão).

---

## 🔧 Pré-requisitos

Antes de começar, garanta que você tenha o Python 3.x instalado em seu sistema.

As seguintes bibliotecas Python são necessárias:

* `pandas`
* `unidecode`
* `openpyxl` (motor de leitura de arquivos `.xlsx`/`.xlsm` para o Pandas)

## ⚙️ Instalação

1.  **Clone o repositório:**
    ```shell
    git clone [https://github.com/BytesAndTracks/consolidate-product-sheets.git](https://github.com/BytesAndTracks/consolidate-product-sheets.git)
    cd consolidate-product-sheets
    ```

2.  **Crie e ative um ambiente virtual (altamente recomendado):**
    ```shell
    # Cria o ambiente
    python -m venv venv

    # Ativa no Windows
    .\venv\Scripts\activate

    # Ativa no macOS/Linux
    source venv/bin/activate
    ```

3.  **Crie um arquivo `requirements.txt`** na raiz do projeto com o seguinte conteúdo:
    ```
    pandas
    unidecode
    openpyxl
    ```

4.  **Instale as dependências:**
    ```shell
    pip install -r requirements.txt
    ```

---

## 🚀 Como Usar

### 1. Processando Dados de Produtos (`processa_produtos.py`)

Este script é configurável através da linha de comando para maior flexibilidade.

**Uso Padrão:**
Execute o script no seu terminal. Ele usará os caminhos e o markup definidos como padrão dentro do código.

```shell
python processa_produtos.py
```

**Uso Avançado (com argumentos):**

* `-i` ou `--input`: Especifica o caminho do arquivo Excel de entrada.
* `-o` ou `--output`: Especifica o nome do arquivo CSV de saída.
* `-m` ou `--markup`: Define o fator multiplicador para o cálculo do preço de venda (ex: `1.55` para 55% de margem).

**Exemplo Prático:**
```shell
python processa_produtos.py -i "C:\Dados\MasterFornecedor.xlsm" -o "produtos_consolidados.csv" -m 1.6
```
Este comando irá ler o arquivo `MasterFornecedor.xlsm`, aplicar um markup de 60% e salvar o resultado em `produtos_consolidados.csv`.

### 2. Processando Dados de Fornecedores (`processa_fornecedores.py`)

Este script possui os caminhos de entrada e saída definidos diretamente no código, sendo ideal para tarefas de rotina com arquivos fixos.

**Passos para Execução:**

1.  **Abra o arquivo** `processa_fornecedores.py` em um editor de código.
2.  **Altere as variáveis** `EXCEL_FILE_PATH` e `OUTPUT_CSV_FILE` para os caminhos corretos em seu sistema.
    ```python
    EXCEL_FILE_PATH = r"C:\Dados\MasterFornecedor.xlsm"
    OUTPUT_CSV_FILE = "fornecedores_para_banco.csv"
    ```
3.  **Execute o script** pelo terminal:
    ```shell
    python processa_fornecedores.py
    ```
O script irá ler a primeira aba da planilha, limpar os dados e gerar o arquivo CSV `fornecedores_para_banco.csv`.

---

## 🤝 Contribuição

Contribuições são bem-vindas! Se você encontrar um bug ou tiver uma sugestão de melhoria, sinta-se à vontade para abrir uma [**Issue**](https://github.com/BytesAndTracks/consolidate-product-sheets/issues) ou enviar um [**Pull Request**](https://github.com/BytesAndTracks/consolidate-product-sheets/pulls).

## 📄 Licença

Este projeto é distribuído sob a licença MIT. Veja o arquivo `LICENSE` para mais detalhes.
