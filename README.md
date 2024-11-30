# Projetos de Web Scraping com Selenium

Este repositório contém vários scripts de web scraping que utilizam o Selenium para coletar dados de diferentes sites. Cada script é projetado para extrair informações específicas de um site e salvar os dados em um arquivo Excel.

## Pré-requisitos

- Python 3.x
- Google Chrome
- [ChromeDriver](https://sites.google.com/chromium.org/driver/) compatível com a versão do seu Chrome
- [Selenium](https://pypi.org/project/selenium/)
- [Pandas](https://pandas.pydata.org/)
- [OpenPyXL](https://openpyxl.readthedocs.io/en/stable/)

## Instalação

1. Crie um ambiente virtual:   ```bash
   python -m venv meu_ambiente_virtual
   source meu_ambiente_virtual/bin/activate   ```

2. Instale as dependências:   ```bash
   pip install -r requirements.txt   ```

## Uso

### 1. FatalModel

- **Descrição**: Coleta nomes e telefones de acompanhantes no site FatalModel.
- **Execução**:  ```bash
  python fatalmodel.py  ```

### 2. Barravipsrio

- **Descrição**: Extrai nomes e números de telefone de acompanhantes no site Barravipsrio.
- **Execução**:  ```bash
  python barravipsrio.py  ```

### 3. GarotaComLocal

- **Descrição**: Coleta dados de acompanhantes no site GarotaComLocal.
- **Execução**:  ```bash
  python garotacomlocal.py  ```

### 4. Lindas

- **Descrição**: Extrai informações de acompanhantes no site Lindas.
- **Execução**:  ```bash
  python lindas.py  ```

### 5. PhotoAcompanhantes

- **Descrição**: Coleta dados de acompanhantes no site PhotoAcompanhantes.
- **Execução**:  ```bash
  python photoacompanhantes.py  ```

### 6. Private55

- **Descrição**: Extrai informações de acompanhantes no site Private55.
- **Execução**:  ```bash
  python private55.py  ```

## Notas

- Certifique-se de que o ChromeDriver está no seu PATH ou no mesmo diretório dos scripts.
- Os scripts são configurados para rodar em modo headless, mas você pode remover essa opção se quiser ver o navegador em ação.
- Os dados coletados são salvos em arquivos Excel no mesmo diretório dos scripts.
