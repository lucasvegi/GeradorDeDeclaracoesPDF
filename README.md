# Gerador de Declarações Personalizadas em PDF 🧾

Este script em Python automatiza a geração de declarações personalizadas em PDF a partir de um template ``.docx`` e de uma planilha ``.xlsx`` contendo os nomes e as matrículas dos participantes.

O funcionamento é simples: ele lê os dados da planilha, substitui os campos ``{{nome}}`` e ``{{matricula}}`` no template ``.docx``, gera os documentos personalizados, converte-os para o formato PDF e organiza os arquivos em diretórios separados por participante.

Este script é especialmente útil para a emissão de declarações de participação em eventos com inscrições feitas por Google Forms. Basta exportar as respostas do formulário em formato ``.xlsx``, adaptá-las conforme o modelo esperado pelo script, e executá-lo para obter todos os certificados ou declarações de forma automatizada.

## Pré-requisitos

1. Windows ou macOS
2. MS Office instalado
3. Instale as seguintes dependências do Python com:

    ```bash
    pip install python-docx openpyxl docx2pdf
    ```

### ℹ️ Nota 1: 
O `docx2pdf` funciona apenas no Windows e macOS. Em Linux, pode ser necessário usar uma alternativa como ``libreoffice`` em modo headless.

### ℹ️ Nota 2: 
No macOS é necessário dar algumas permissões no sistema para o script funcionar adequadamente.

1. Vá em Preferências do Sistema → Segurança e Privacidade → aba Privacidade.
2. Clique no cadeado no canto inferior esquerdo e digite sua senha/admin.
3. Selecione "Acesso Total ao Disco" na barra lateral.
4. Encontre o Microsoft Word na lista e conceda essa permissão.

## Como usar
1. Prepare um arquivo Word (``modelo_declaracao.docx``) com os placeholders ``{{nome}}`` e ``{{matricula}}``.
2. Crie uma planilha (``alunos.xlsx``) com duas colunas:
    - Coluna A: Matrícula
    - Coluna B: Nome
    - **OBS:** a primeira linha da planilha deve ser de cabeçalho, com os dados dos alunos iniciando a partir da segunda linha.
3. Coloque esses arquivos no mesmo diretório do script e execute:

    ```bash
    python GeradorDeclaracoes.py
    ```

### ℹ️ Nota: 
Altere os caminhos das variáveis ``template_path``, ``planilha_path`` e ``saida_dir`` no final do script conforme a sua necessidade.

## Comportamento da execução do script

### 📈 Barra de Progresso

O script mostra uma barra de progresso durante a execução, indicando quantos PDFs foram gerados.

### 📁 Saída
- Arquivos ``.docx`` são salvos na pasta ``/declaracoes/docx/``
- Arquivos ``.pdf`` são salvos na pasta ``/declaracoes/pdf/``
