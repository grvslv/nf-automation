# Automação de Notas Fiscais

Este repositório contém um script Python que automatiza o processamento de planilhas e a organização de arquivos de notas fiscais. Ele lê planilhas de protocolo, extrai as chaves de acesso, copia os arquivos XML correspondentes e move as DANFEs em PDF para pastas de destino.

## Instalação

1. Certifique-se de ter o Python 3 instalado no Windows.
2. Instale as dependências necessárias (pandas, openpyxl, tqdm). Uma forma simples é:
   ```cmd
   pip install -r requirements.txt
   ```
   Se preferir, você também pode executar:
   ```cmd
   pip install pandas openpyxl tqdm
   ```
   O módulo `tkinter` já acompanha o Python padrão no Windows.

## Execução

No prompt de comando do Windows, navegue até a pasta do repositório e execute:

```cmd
python Downloads\nfe_bot_publico.py
```

Os caminhos de entrada e saída estão definidos no início do script e podem ser ajustados conforme a necessidade.
