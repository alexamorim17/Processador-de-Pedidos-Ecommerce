import pandas as pd
import tkinter as tk
from tkinter import filedialog
import re
from pathlib import Path
from datetime import datetime

def consultarTamanho(texto):

    tamanho = ["PP", "P", "M", "G", "GG", "XG", "G1", "G2", "G3"]

    return next(
        (t for t in sorted(tamanho, key=len, reverse=True) if t in texto),
        None
    )
    

def extrair_nick(texto):
    if not isinstance(texto, str):
        return None

    # Procura o padrão [qualquer_coisa:nick]
    match = re.search(r"\[[^:\]]+:(.*?)\]", texto)
    if match:
        return match.group(1)  # retorna só o nick
    return None

def abrir_pasta_lista():
    return Path.cwd().joinpath("lista.xlsx").resolve()



def extrair_nome(texto):
    
    pasta_atual = abrir_pasta_lista()
    lista_tratamento = pd.read_excel(pasta_atual, sheet_name="lista")

    lista_planilha = pd.DataFrame(lista_tratamento)



    return next(
    (
        nome_tratado
        for nome, nome_tratado in zip(
            lista_planilha['Nome'],
            lista_planilha['nome_tratado']
        )
        if nome in texto
    ),
    None
)



def converter_data(data):
    # Remove o "Z" no final se existir
    if isinstance(data, str) and data.endswith("Z"):
        data = data[:-1]
    
    # Converte considerando a hora (mesmo que não usemos)
    data_dt = datetime.strptime(data, "%Y-%m-%d %H:%M:%S")
    
    return data_dt.strftime("%d/%m/%Y")



root = tk.Tk()
root.withdraw()


file_path = filedialog.askopenfilename(
    title="Selecione um arquivo",
    filetypes=(
        ("Arquivos Excel", "*.xlsx *.xls"),
        ("Todos os arquivos", "*.*")
    )
)

x = pd.read_excel(file_path, sheet_name= "Report")




col = ["Creation Date", "Order", "Client Name", "Client Last Name", "SLA Type", "SKU Name","Item Attachments","TAMANHO","NICK","PATCH", "Estimate Delivery Date"]

z = pd.DataFrame(x,columns= col)
z["TAMANHO"] = z["SKU Name"].apply(consultarTamanho)
z["NICK"] = z["Item Attachments"].apply(extrair_nick)
z["PATCH"] = z["SKU Name"].apply(extrair_nome)
z["CLIENTE"] = z["Client Name"].astype(str) + " " + z["Client Last Name"].astype(str)
z["Estimate Delivery Date"] = z["Estimate Delivery Date"].apply(converter_data)
z["Creation Date"] = z["Creation Date"].apply(converter_data)

z = z.drop(columns=["SKU Name", "Item Attachments","Client Name", "Client Last Name"])

z = z.rename(columns={
    "Creation Date": "DATA",
    "Order": "PEDIDO",
    "Client Name": "CLIENTE",
    "SLA Type": "ENVIO",
    "TAMANHO":"TAM",
    "Estimate Delivery Date": "OBS"
})

print(z)

z = z[["DATA", "PEDIDO", "CLIENTE", "ENVIO", "TAM", "NICK", "PATCH", "OBS"]]


with pd.ExcelWriter(
    'planilha_nova.xlsx',
    engine='openpyxl',
    mode='a',                   # modo append
    if_sheet_exists='replace'   # substitui a aba se já existir
) as writer:
    z.to_excel(writer, sheet_name='planilha_nova', index=False)





