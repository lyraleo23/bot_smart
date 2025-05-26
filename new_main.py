import os
import time
import requests
import pandas as pd

import tkinter as tk
from tkinter import filedialog, ttk
from tkcalendar import Calendar

from smartphar import open_smartphar, login_smartphar, open_receitas_screen

# Legado
import ajuste_dataframe
from smartphar import insert_orders_smartphar


def select_folder(folder_label):
    folder_selected = filedialog.askdirectory()
    folder_label.config(text=folder_selected)
    return folder_selected


def merge_excel_files(folder_path, total_pedidos_label, total_requisicoes_label):
    print(folder_path)
    if not folder_path:
        print('Error: No folder selected')
        return
    else:
        # Merge excel files
        print('Merging excel files...')
        df_list = []
        output_file = "pedidos.xlsx"

        if output_file in os.listdir(folder_path):
            os.remove(folder_path + "/" + output_file)

        if 'filtered_orders.xlsx' in os.listdir(folder_path):
            os.remove(folder_path + "/" + 'filtered_orders.xlsx')

        if 'pedidos.xlsx' in os.listdir(folder_path):
            os.remove(folder_path + "/" + 'pedidos.xlsx')

        for file in os.listdir(folder_path):
            if file.endswith('.xlsx') or file.endswith('.xls'):
                df = pd.read_excel(os.path.join(folder_path, file))
                df_list.append(df)
        combined_df = pd.concat(df_list)
        
        combined_df.to_excel(folder_path + "/" + output_file, index=False)
        print('Files merged!')

        unmanufactured_products = get_unmanufactured_products()    
        total_pedidos, total_requisicoes = build_dataframe(folder_path, unmanufactured_products)

        def update_production_total():
            total_pedidos_label.config(text=f'Total pedidos: {total_pedidos}')
            total_requisicoes_label.config(text=f'Total reqs: {total_requisicoes}')

        update_production_total()
        return


def get_unmanufactured_products():
    response_semi = requests.get("https://api.fmiligrama.com/produtos?tipoInterno=Semi-Acabado").json()
    response_semi = pd.json_normalize(response_semi['data'])
    response_industrializados = requests.get("https://api.fmiligrama.com/produtos?tipoInterno=Industrializado").json()
    response_industrializados = pd.json_normalize(response_industrializados['data'])
    return pd.concat([response_semi, response_industrializados])


def build_dataframe(folder_path, unmanufactured_products):
    #Cria um Data Frame com a planilha de pedidos
    orders = pd.read_excel(folder_path + "/" + 'pedidos.xlsx', sheet_name='Sheet1')
    orders.fillna({'Fone': '99999999999'}, inplace= True)
    orders.fillna({'Celular': '99999999999'}, inplace= True)

    #Agrupa os itens duplicados
    orders = orders.groupby(["Número do pedido","Fone", "CPF/CNPJ", "Nome do contato", "Código (SKU)"]).sum("Quantidade")
    orders = orders.reset_index()

    #Converte a coluna codigo da base de dados para inteiro
    unmanufactured_products = unmanufactured_products[unmanufactured_products['codigo'] != 'folder_sac']
    unmanufactured_products["codigo"] = unmanufactured_products["codigo"].astype(int)

    #Mescla a planilha de pedidos com a base de dados para filtrar os pedidos semi-acabados
    filtered_orders = pd.merge(orders, unmanufactured_products, left_on="Código (SKU)", right_on="codigo", how="left")
    filtered_orders = filtered_orders.sort_values(by='Número do pedido', ascending=True)

    #Ajusta os dados do dataframe
    ajuste_dataframe.ajuste_excel(filtered_orders)

    # #Armazena os pedidos manipulados
    filtered_orders["tipoInterno"] = filtered_orders["tipoInterno"].str.strip()
    filtered_orders["tipoInterno"] = filtered_orders["tipoInterno"].str.lower()
    manipulados = filtered_orders[filtered_orders["tipoInterno"].isin(["manipulado"])]

    #Soma a quantidade de reqs e pedidos
    total_pedidos = manipulados["Número do pedido"].drop_duplicates()
    total_pedidos = total_pedidos.count()
    total_requisicoes = manipulados["Quantidade"].sum()

    print(f'Total Pedidos: {total_pedidos}')
    print(f'Total Requisições: {total_requisicoes}')

    filtered_orders.to_excel(folder_path + "/" + 'filtered_orders.xlsx', index=False)
    return total_pedidos, total_requisicoes


def filter_manipulados(folder_path):
    print('filter_manipulados')
    # read filtered_orders.xslx file
    filtered_orders = pd.read_excel(folder_path + "/" + 'filtered_orders.xlsx', sheet_name='Sheet1')

    # Filter only the orders that need to be included in smartphar
    smart_filtered_orders = filtered_orders["tipoInterno"] == "manipulado"
    smart_filtered_orders = filtered_orders.loc[smart_filtered_orders]

    return smart_filtered_orders


def include_reqs(folder_path, sector_var, production_branch, production_date):
    # Filter orders
    smart_filtered_orders = filter_manipulados(folder_path)
    
    # Open and login to Smartphar
    open_smartphar()
    login_smartphar(production_branch, sector_var)
    time.sleep(2)

    # Open receitas screen
    open_receitas_screen()
    time.sleep(2)

    # Insert orders
    insert_orders_smartphar(smart_filtered_orders, sector_var, production_branch, production_date)

    # Create a new window to indicate completion
    completed_window = tk.Toplevel()
    completed_window.title("Status")
    completed_window.attributes('-topmost', True)  # Make the window appear above all others
    completed_label = tk.Label(completed_window, text="CONCLUIDO", font=("Helvetica", 70))
    completed_label.pack(pady=20)
    ok_button = tk.Button(completed_window, text="OK", command=completed_window.destroy)
    ok_button.pack(pady=10)


def main():
    root = tk.Tk()
    root.title("Smartphar Hub")

    # Folder selection
    folder_label = tk.Label(root, text="No folder selected")
    folder_label.pack(pady=5)
    folder_path = tk.StringVar()
    folder_button = tk.Button(root, text="Select Folder", command=lambda: folder_path.set(select_folder(folder_label)))
    folder_button.pack(pady=5)

    # Merge button
    merge_button = tk.Button(root, text="Merge Excel Files", command=lambda: merge_excel_files(folder_path.get(), total_pedidos_label, total_requisicoes_label))
    merge_button.pack(pady=20)

    # Date selector
    cal_label = tk.Label(root, text="Select Date")
    cal_label.pack(pady=5)
    cal = Calendar(root, selectmode='day')
    cal.pack(pady=5)
    production_date = tk.StringVar(value=cal.get_date())

    def update_production_date(event):
        production_date.set(cal.get_date())
        date = production_date.get().split('/')
        month = date[0]
        day = date[1]
        year = date[2]

        if len(day) == 1:
            day = '0' + day
        if len(month) == 1:
            month = '0' + month
        if len(year) == 2:
            year = '20' + year

        production_date.set(f'{day}{month}{year}')
        print(production_date.get())
    cal.bind("<<CalendarSelected>>", update_production_date)

    #Rótulos de quantidades de pedidos com base na analise dos pedidos
    total_pedidos_label = tk.Label(root, text="Total pedidos: ")
    total_pedidos_label.pack(pady=5)

    #Rótulos de quantidades de reqs com base na analise dos pedidos
    total_requisicoes_label = tk.Label(root, text="Total reqs: ")
    total_requisicoes_label.pack(pady=5)
    
    # Setor selector
    sector_label = tk.Label(root, text="Selecione o setor")
    sector_label.pack(pady=5)
    sector_var = tk.StringVar(value="Site")
    sector_menu = ttk.Combobox(root, textvariable=sector_var, values=["Site", "Manual"])
    sector_menu.pack(pady=5)

    # Setor filial de produção
    production_branch_label = tk.Label(root, text="Selecione a filial de produção")
    production_branch_label.pack(pady=5)
    production_branch_var = tk.StringVar(value="100")
    production_branch_menu = ttk.Combobox(root, textvariable=production_branch_var, values=["100", "600"])
    production_branch_menu.pack(pady=5)

    # Run button
    merge_button = tk.Button(root, text="Run", command=lambda: include_reqs(folder_path.get(), sector_var.get(), production_branch_var.get(), production_date.get()))
    merge_button.pack(pady=20)
    
    root.mainloop()

main()
