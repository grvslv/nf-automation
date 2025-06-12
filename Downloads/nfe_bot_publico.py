
import pandas as pd
import os
import shutil
import tkinter as tk
from tkinter import messagebox
from tqdm import tqdm
import re

# Caminhos das pastas (exemplos fictícios)
pasta_excel = r"C:\Automacao\Protocolos"
pasta_xml = r"D:\Automacao\XML\NF-es"
pasta_resultado = r"C:\Automacao\Xml"
pasta_danfes = r"D:\Automacao\DANFES"
pasta_resultado_pdf = r"C:\Automacao\XmlPdf"
pasta_logs = os.path.join(pasta_excel, "Logs")
arquivo_txt = os.path.join(pasta_logs, "chaves_invalidas.txt")

# Garantir pastas
os.makedirs(pasta_logs, exist_ok=True)
os.makedirs(pasta_resultado, exist_ok=True)
os.makedirs(pasta_resultado_pdf, exist_ok=True)

def reiniciar_log():
    with open(arquivo_txt, "w", encoding="utf-8") as f:
        pass

def mostrar_alerta(mensagens):
    root = tk.Tk()
    root.withdraw()
    messagebox.showwarning("Problemas encontrados", "\n".join(mensagens))
    root.destroy()

def mostrar_janela_carregando():
    root = tk.Tk()
    root.withdraw()
    loading = tk.Toplevel()
    loading.title("Processando")
    loading.geometry("320x80")
    label = tk.Label(
        loading,
        text="Processando arquivos... Aguarde...",
        font=("Arial", 10),
    )
    label.pack(pady=25)
    loading.update()
    return loading

def encontrar_linha_cabecalho(arquivo):
    df_temp = pd.read_excel(arquivo, engine="openpyxl", header=None)
    for i, row in df_temp.iterrows():
        if row.astype(str).str.contains("Chave de acesso", case=False, na=False).any():
            return i
    return None

def valida_chave(chave):
    chave = re.sub(r"\D", "", chave)
    return chave if len(chave) == 44 else None

def indexar_arquivos(diretorio, extensao):
    index = {}
    for root, dirs, files in os.walk(diretorio):
        for f in files:
            if f.lower().endswith(extensao):
                nome = os.path.splitext(f)[0]
                index[nome] = os.path.join(root, f)
    return index

# Início da execução
reiniciar_log()
mensagens_alerta = []

# Indexação única
print(" Indexando arquivos XML e PDF...")
index_xml = indexar_arquivos(pasta_xml, ".xml")
index_pdf = indexar_arquivos(pasta_danfes, ".pdf")

arquivos_excel = [os.path.join(pasta_excel, a) for a in os.listdir(pasta_excel)
                  if a.endswith(('.xls', '.xlsx')) and not a.startswith('~$')]

janela = mostrar_janela_carregando()

for arquivo in arquivos_excel:
    print(f"Processando: {arquivo}")
    try:
        linha_cabecalho = encontrar_linha_cabecalho(arquivo)
        if linha_cabecalho is None:
            print(" Cabeçalho não encontrado.")
            continue

        df = pd.read_excel(arquivo, engine="openpyxl", header=linha_cabecalho)
        coluna_chave = next((c for c in df.columns if "chave de acesso" in str(c).lower()), None)
        if not coluna_chave:
            print(" Coluna 'Chave de acesso' não encontrada.")
            continue

        chaves_cruas = df[coluna_chave].dropna().astype(str)
        chaves_processadas = [valida_chave(ch) for ch in chaves_cruas]
        chaves_validas = [ch for ch in chaves_processadas if ch]
        chaves_invalidas = [ch for ch in chaves_cruas if not valida_chave(ch)]

        if chaves_invalidas:
            mensagens_alerta += chaves_invalidas
            with open(arquivo_txt, "a", encoding="utf-8") as f:
                f.write(f"\nChaves inválidas no arquivo {os.path.basename(arquivo)}:\n")
                for chave in chaves_invalidas:
                    f.write(f"{chave}\n")
            continue

        errors_found = False

        for chave in tqdm(chaves_validas, desc="Processando arquivos"):

            # Copiar XML
            if chave in index_xml:
                shutil.copy(index_xml[chave], pasta_resultado)
                print(f"{chave}.xml copiado.")
            else:
                mensagens_alerta.append(f"{chave} - XML não encontrado")
                with open(arquivo_txt, "a", encoding="utf-8") as f:
                    f.write(f"{chave} - XML não encontrado\n")
                errors_found = True

            # Mover PDF
            if chave in index_pdf:
                shutil.move(index_pdf[chave], pasta_resultado_pdf)
                print(f"{chave}.pdf movido.")
            else:
                mensagens_alerta.append(f"{chave} - PDF não encontrado")
                with open(arquivo_txt, "a", encoding="utf-8") as f:
                    f.write(f"{chave} - PDF não encontrado\n")
                errors_found = True

        if not errors_found:
            os.remove(arquivo)
            print(f"{arquivo} removido após processamento.")
        else:
            print(f"{arquivo} NÃO foi removido devido a erros.")

    except Exception as e:
        print(f"Erro ao processar {arquivo}: {e}")

janela.destroy()

if mensagens_alerta:
    mostrar_alerta(mensagens_alerta)
else:
    if os.path.exists(arquivo_txt):
        os.remove(arquivo_txt)
    print("Tudo processado sem erros.")
