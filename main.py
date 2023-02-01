import datetime as dt
import tkinter as tk
from tkinter import ttk
import pandas as pd

materiais = pd.read_excel('materiais.xlsx', engine='openpyxl')

lista_itens = []
lista_tipos = ['Galão', 'Caixa', 'Saco', 'Unidade']

janela = tk.Tk()

def salvar_item():
    descricao = entry_descricao.get()
    tipo = combobox_selecionar_tipo.get()
    quant = entry_quant.get()
    data_criacao = dt.datetime.now()
    data_criacao = data_criacao.strftime("%d/%m/%Y %H:%M")
    itens = materiais.shape[0] + len(lista_itens)+1
    itens_str = "ID-{}".format(itens)
    lista_itens.append((itens_str, descricao, tipo, quant, data_criacao))

#CRIAÇÃO DA JANELA

janela.title('Cadastro de materias')

label_descricao = tk.Label(text='Item:')
label_descricao.grid(row=1, column=0, padx=10, pady=10, sticky='nswe', columnspan=4)

entry_descricao = tk.Entry()
entry_descricao.grid(row=2, column=0, padx=10, pady=10, sticky='nswe', columnspan=4)

label_tipo_unidade = tk.Label(text='Tipo da unidade do item:')
label_tipo_unidade.grid(row=3, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

combobox_selecionar_tipo = ttk.Combobox(values=lista_tipos)
combobox_selecionar_tipo.grid(row=3, column=2, padx=10, pady=10, sticky='nswe', columnspan=2)

label_quant = tk.Label(text='Quantidade do tipo de unidade:')
label_quant.grid(row=4, column=0, padx=10, pady=10, sticky='nswe', columnspan=2)

entry_quant = tk.Entry()
entry_quant.grid(row=4, column=2, padx=10, pady=10, sticky='nswe', columnspan=2)

botao_salvar_item = tk.Button(text='Salvar item', command=salvar_item)
botao_salvar_item.grid(row=5, column=0, padx=10, pady=10, sticky='nswe', columnspan=4)

janela.mainloop()

novo_material = pd.DataFrame(lista_itens, columns=['Item', 'Descrição', 'Tipo', 'Quantidade', 'Data de Criação'])
materiais = materiais.append(novo_material, ignore_index=True)
materiais.to_excel('materiais.xlsx', index=False)