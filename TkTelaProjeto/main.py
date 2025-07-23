import tkinter as tk
from tkinter import ttk
from tkcalendar import DateEntry
from tkinter.filedialog import askopenfilename
import requests
import pandas as pd
from datetime import datetime
import numpy as np


requisicao = requests.get('https://economia.awesomeapi.com.br/json/all')
dic_moedas = requisicao.json()

lista_moedas_temp = list(dic_moedas.keys())


def pegar_cotacao():
    moeda = combobox_moeda.get()
    data = calendario_moeda.get()
    ano = data[-4:]
    mes = data [3:5]
    dia = data [:2]
    link = f'https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano}{mes}{dia}&end_date={ano}{mes}{dia}'
    requisicao_moeda = requests.get(link)
    cotacao = requisicao_moeda.json() #Transformou o json em lista Python
    valor_moeda = cotacao[0]["bid"]
    label_texto_cotacao["text"] = f'A cotação {moeda} no dia {dia}/{mes}/{ano} está no valor de R${valor_moeda}'

def buscar_arquivo():
    caminho_arquivo = askopenfilename(title="Selecione o arquivo de Moeda")
    variavel_caminho_arquivo.set(caminho_arquivo)
    if caminho_arquivo:
        label_arquivo_selecionado["text"] =  f'Arquivo Selcionado: {caminho_arquivo}'


def atualizar_cotacoes():
    try:
        df = pd.read_excel(variavel_caminho_arquivo.get())
        moedas = df.iloc[:, 0]

        data_inicial = calendario_data_inicial.get()  # 17/06/2025
        data_final = calendario_data_final.get()

        dt_inicio = datetime.strptime(data_inicial, "%d/%m/%Y")
        dt_final = datetime.strptime(data_final, "%d/%m/%Y")

        # Corrigido: quantidade de dias corretamente (inclusivo)
        quantidade_dias = (dt_final - dt_inicio).days + 1

        ano_i, mes_i, dia_i = data_inicial[-4:], data_inicial[3:5], data_inicial[:2]
        ano_f, mes_f, dia_f = data_final[-4:], data_final[3:5], data_final[:2]

        for moeda in moedas:
            url = (f"https://economia.awesomeapi.com.br/json/daily/{moeda}-BRL/?start_date={ano_i}{mes_i}{dia_i}&end_date={ano_f}{mes_f}{dia_f}")
            requisicao = requests.get(url)
            cotacoes = requisicao.json()

            for cotacao in cotacoes:
                data_str = datetime.strptime(
                    cotacao["create_date"], "%Y-%m-%d %H:%M:%S"
                ).strftime("%d/%m/%Y")

                bid = float(cotacao["bid"])

                # se ainda não existe coluna para esse dia, cria‑a
                if data_str not in df.columns:
                    df[data_str] = np.nan

                # preenche a linha da moeda
                df.loc[df.iloc[:, 0] == moeda, data_str] = bid

        # salva de uma vez só
        #df.to_excel("Cotações Text.xlsx", index=False)
        label_atualizar_cotacoes["text"] = "Arquivo Atualizado com Sucesso"
        print(ano_i, mes_i, dia_i) # ESTÁ COM 2 ERROS
        print(ano_f, mes_f, dia_f)
        print(url)

    except:
        label_atualizar_cotacoes["text"] = "Selecione um Arquivo Excel no Formato Correto"

janela = tk.Tk()

janela.title("Ferramenta de Cotação de Moedas")

titulo1 = tk.Label(text="Cotação de 1 moeda específica:", borderwidth=2, relief="solid")
titulo1.grid(row=0, column=0, columnspan=3, sticky="news", padx=10, pady=10)

label_selecionar_moeda = tk.Label(text="Selecionar Moeda Que Deseja Consultar", anchor="e")
label_selecionar_moeda.grid(row=1, column=0, columnspan=2, sticky="news", padx=10, pady=10)

#Combobox para selecionar a moeda
combobox_moeda = ttk.Combobox(values=lista_moedas_temp)
combobox_moeda.grid(row=1, column=2, sticky="news", padx=10, pady=10)

label_selecionar_dia = tk.Label(text="Qual Dia Deseja Pegar Cotação", anchor="e")
label_selecionar_dia.grid(row=2, column=0, columnspan=2, sticky="news", padx=10, pady=10)

calendario_moeda = DateEntry(year=2025, locale='pt_BR')
calendario_moeda.grid(row=2, column=2, sticky="news", padx=10, pady=10)

label_texto_cotacao = tk.Label(text=" ")
label_texto_cotacao.grid(row=3, column=0, columnspan=2, sticky="news", padx=10, pady=10)

botao_pegar_cotacao = tk.Button(text="Pegar Cotação", command=pegar_cotacao)
botao_pegar_cotacao.grid(row=3, column=2, sticky="news", padx=10, pady=10)


# -*-*-*-*-*-*-*-*-*-*-*-*
titulo2 = tk.Label(text="Cotacao de multiplas moedas:", borderwidth=2, relief="solid")
titulo2.grid(row=4, column=0, columnspan=3, sticky="news", padx=10, pady=10)

label_selecionar_arquivos = tk.Label(text="Selecione um Arquivo Excel com as Moedas na Coluna 'A'",)
label_selecionar_arquivos.grid(row=5, column=0, columnspan=2, sticky="news", padx=10, pady=10)

variavel_caminho_arquivo = tk.StringVar()

botao_buscar_arquivos = tk.Button(text="Clique para Selecionar", command=buscar_arquivo)
botao_buscar_arquivos.grid(row=5, column=2, sticky="news", padx=10, pady=10)

label_arquivo_selecionado = tk.Label(text="Nenhum Arquivo Selecionado", anchor="e")
label_arquivo_selecionado.grid(row=6, column=0, columnspan=3, sticky="news", padx=10, pady=10)

label_data_inicial = tk.Label(text="A data Inicial:", anchor="e")
label_data_inicial.grid(row=7, column=0, sticky="news", padx=10, pady=10)

calendario_data_inicial = DateEntry(year=2025, locale='pt_br')
calendario_data_inicial.grid(row=7, column=1, sticky="news", padx=10, pady=10)

label_data_final = tk.Label(text="A data Final:", anchor="e")
label_data_final.grid(row=8, column=0, sticky="news", padx=10, pady=10)

calendario_data_final = DateEntry(year=2025, locale='pt_br')
calendario_data_final.grid(row=8, column=1, sticky="news", padx=10, pady=10)

botao_atualizar_cotacoes = tk.Button(text="Atualizar Cotações", command=atualizar_cotacoes)
botao_atualizar_cotacoes.grid(row=9, column=0, sticky="news", padx=10, pady=10)

label_atualizar_cotacoes = tk.Label(text=" ", anchor="e")
label_atualizar_cotacoes.grid(row=9, column=1, columnspan=2, sticky="news", padx=10, pady=10)

botao_fechar = tk.Button(text="Fechar", command=janela.quit)
botao_fechar.grid(row=10, column=2, sticky="news", padx=10, pady=10)

janela.mainloop()