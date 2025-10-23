# Gerador de Enum Java a partir de XLSX

Este é um script Python para automatizar a criação de classes `Enum` em Java a partir de dados em uma planilha Excel (`.xlsx`).

## Recursos

* **Leitura de `.xlsx`**: Lê dados diretamente de uma planilha Excel.
* **Mapeamento de Colunas**: Permite configurar facilmente quais colunas da planilha correspondem a quais atributos do Enum.
* **Conversão de Tipos**: Converte automaticamente os dados da planilha:
    * Campos com prefixo `ind_` ou `indicador` tornam-se `boolean` (tratando `1`, `S`, `V`, `N/A`, etc.).
    * Campos com prefixo `perc` tornam-se `double`.
    * Outros campos tornam-se `String` (formatados com aspas).
* **Ignora Linhas Riscadas**: O script verifica uma coluna específica e ignora qualquer linha onde o texto dessa célula esteja formatado como "riscado" (strikethrough).
* **Geração de Código**: Cria um arquivo `.java` completo, com atributos, construtor e getters.

## Como Usar

### 1. Pré-requisitos

* Python 3.x
* O módulo `venv` do Python (em sistemas Debian/Ubuntu, instale com `sudo apt install python3.xx-venv`)

### 2. Instalação

1.  Clone este repositório:

    ```bash
    git clone <url-do-seu-repositorio>
    cd <nome-do-repositorio>
    ```

2.  Crie e ative um ambiente virtual:

    ```bash
    # Criar o ambiente
    python3 -m venv .venv
    
    # Ativar (Linux/macOS)
    source .venv/bin/activate
    
    # Ativar (Windows PowerShell)
    # .\.venv\Scripts\Activate.ps1
    ```

3.  Instale as dependências:

    ```bash
    pip install pandas openpyxl
    ```

### 3. Configuração

Abra o arquivo `gerador_enum.py` e edite a seção `CONFIGURAÇÕES` no topo do arquivo:

* `ARQUIVO_EXCEL`: O nome do seu arquivo Excel (ex: `'minha_tabela.xlsx'`).
* `NOME_DA_ABA`: O nome exato da aba (sheet) dentro da planilha (ex: `'Plan1'`).
* `NOME_DO_ENUM`: O nome que você deseja para a sua classe Enum (ex: `'MeuSuperEnum'`).
* `MAPEAMENTO_COLUNAS`: A parte mais importante. Mapeie os nomes dos atributos Java (chave) para os nomes exatos das colunas no Excel (valor).
* `COLUNA_VERIFICACAO_RISCADO`: A coluna que o script deve checar para pular linhas (ex: `'id_codigo'`).

### 4. Execução

Com seu ambiente virtual ativado e as dependências instaladas, basta executar o script:

```bash
python gerador_enum.py