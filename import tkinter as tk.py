import tkinter as tk
from tkinter import ttk, messagebox
import pyodbc
from tabulate import tabulate

SERVER = '192.168.0.254'
DATABASE = 'BDCOPIA'
USERNAME = 'microuni'
PASSWORD = 'microuni'

# Função para conectar e buscar os dados da nota fiscal
def buscar_nf():
    numnota = entry_nota.get()
    if not numnota:
        messagebox.showwarning("Aviso", "Por favor, insira um número de nota fiscal.")
        return

    conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+SERVER+';DATABASE='+DATABASE+';UID='+USERNAME+';PWD='+PASSWORD)
    cursor = conn.cursor()

    cursor.execute("""
        WITH calculateddata AS (
        SELECT
            convert(varchar, b.Dtultrea, 3) 'REM',
            convert(varchar, b.DtUltComp, 3) 'ENT',
            a.codpro COD,
            b.descr DESCRICAO,
            FORMAT(a.quant, '0.00') QUAN,
            a.unidade UND,
            FORMAT(a.preco, '0.00') RS_LISTA,
            FORMAT(a.custven, '0.00') RS_CUSTO,
            FORMAT(b.precoven, '0.00') RS_VEN_REAL,
            CASE WHEN b.margemluc=0 THEN 0 
            ELSE CAST(((b.margemluc/(100-b.margemluc))*100) AS DECIMAL(7,1)) 
            END AS "MKP",
            CASE WHEN b.precocomp=0 THEN 0
            ELSE CAST(((b.precoven-b.precocomp)/b.precocomp)*100 AS DECIMAL(7,1))
            END AS MKP_REAL,
            FORMAT((a.valsubstri)/a.quant, '0.00') ST
        FROM
            ITNFENTCAD a,
            produtocad b,
            complementoproduto c
        WHERE
            a.codpro = b.codpro AND
            b.codpro = c.codpro AND
            a.numnota = ?
        )
        SELECT
            [ENT],
            [REM],
            [COD],
            [DESCRICAO],
            [QUAN],
            [UND],
            [RS_LISTA],
            [RS_CUSTO],
            FORMAT([RS_CUSTO] * (1+([MKP]/100)), '0.00') AS [RS_VENDA_SUG],
            [RS_VEN_REAL],
            [MKP_REAL],
            [MKP],
            [MKP_REAL] - [MKP] AS [DIF_MKP],
            [ST]
        FROM calculateddata
    """, (numnota,))
    
    records = cursor.fetchall()
    conn.close()

    if not records:
        messagebox.showwarning("Aviso", "Nota Fiscal não encontrada.")
    else:
        # Limpar a tabela antes de exibir os novos dados
        for row in tree.get_children():
            tree.delete(row)

        # Adicionar os dados na Treeview
        for i, r in enumerate(records):
            tree.insert('', 'end', values=(r.ENT, r.REM, r.COD, r.DESCRICAO, r.QUAN, r.UND, r.RS_LISTA, r.RS_CUSTO, r.RS_VENDA_SUG, r.RS_VEN_REAL))

# Função para remarcar os itens selecionados
def remarcar_selecionados():
    selected_items = tree.selection()
    if not selected_items:
        messagebox.showwarning("Aviso", "Por favor, selecione os itens que deseja remarcar.")
        return

    codigos_selecionados = [tree.item(item, 'values')[2] for item in selected_items]  # Pegar os códigos das linhas selecionadas

    conn = pyodbc.connect('DRIVER={SQL Server};SERVER='+SERVER+';DATABASE='+DATABASE+';UID='+USERNAME+';PWD='+PASSWORD)
    cursor = conn.cursor()

    # Atualizar o preço sugerido para os códigos selecionados
    for item in selected_items:
        r_sug = tree.item(item, 'values')[8]  # Valor de RS_VENDA_SUG
        codpro = tree.item(item, 'values')[2]  # Código do produto
        cursor.execute(f"""
            UPDATE produtocad
            SET precoven = {r_sug}
            WHERE codpro = {codpro}
        """)

    conn.commit()
    conn.close()

    messagebox.showinfo("Sucesso", "Itens remarcados com sucesso!")

# Configuração da janela principal
root = tk.Tk()
root.title("Remarcar Itens")

# Configuração dos widgets
frame = tk.Frame(root)
frame.pack(pady=10)

label_nota = tk.Label(frame, text="Digite o Nº da NF:")
label_nota.grid(row=0, column=0)

entry_nota = tk.Entry(frame)
entry_nota.grid(row=0, column=1)

btn_buscar = tk.Button(frame, text="Buscar", command=buscar_nf)
btn_buscar.grid(row=0, column=2, padx=10)

# Configuração da tabela (Treeview)
columns = ("ENT", "REM", "COD", "DESCRICAO", "QUAN", "UND", "RS_LISTA", "RS_CUSTO", "RS_VENDA_SUG", "RS_VEN_REAL")
tree = ttk.Treeview(root, columns=columns, show="headings", height=10)

for col in columns:
    tree.heading(col, text=col)
    tree.column(col, width=100)

tree.pack(pady=20)

# Botão para remarcar itens selecionados
btn_remarcar = tk.Button(root, text="Remarcar Selecionados", command=remarcar_selecionados)
btn_remarcar.pack(pady=10)

# Loop da interface gráfica
root.mainloop()
