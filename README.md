# Gerenciador-de-Estoque
#Gerenciador de estoque feito para a Coordenadoria de Saúde e Assistência Social e Religiosa da PMCE

from tkinter import *
from tkinter import ttk, Tk
import sqlite3
import pandas as pd
from openpyxl.workbook import Workbook

from babel import numbers
from tkcalendar import Calendar, DateEntry
import datetime


janela = Tk()


class funcs():

    def exportar_excel(self):
        self.conectar_banco()
        self.cursor.execute("""SELECT * FROM estoque""")
        estoque_cadastrado = self.cursor.fetchall()
        estoque_cadastrado = pd.DataFrame(estoque_cadastrado,
                                          columns=['Código', 'Nota Fiscal', 'Código Produto', 'Descrição', 'Data',
                                                   'Unidade', 'Quantidade Recebida', 'Saida de Material',
                                                   'Valor Unitário', 'Valor Total', 'Quantidade no Estoque'])

        # Obter a data atual no formato DD.MM.YYYY
        data_atual = datetime.datetime.now().strftime("%d.%m.%Y")

        # Adicionar a data ao nome do arquivo
        nome_arquivo = f'Estoque_CSASR_{data_atual}.xlsx'

        # Salvar o DataFrame como um arquivo Excel com o nome atualizado
        estoque_cadastrado.to_excel(nome_arquivo, index=False)

        self.conn.commit()
        self.desconectar_banco()

    def conectar_banco(self):
        self.conn = sqlite3.connect("CSARS.bd")
        self.cursor = self.conn.cursor()

    def desconectar_banco(self):
        self.conn.close()

    def limpar_a_tela(self):
        self.entrada_codigo.delete(0, END)
        self.entrada_notafiscal.delete(0, END)
        self.entrada_codproduto.delete(0, END)
        self.entrada_descricao.delete(0, END)
        self.entrada_data.delete(0, END)
        self.entrada_unidade.delete(0, END)
        self.entrada_quantidade_recebida.delete(0, END)
        self.entrada_saida_material.delete(0, END)
        self.entrada_valor_unitario.delete(0, END)
        self.entrada_valor_total.delete(0, END)

    def variaveis(self):
        self.codigo = self.entrada_codigo.get()
        self.notafiscal = self.entrada_notafiscal.get()
        self.codproduto = self.entrada_codproduto.get()
        self.descricao = self.entrada_descricao.get()
        self.data = self.entrada_data.get()
        self.unidade = self.entrada_unidade.get()
        self.quantidaderecebida = self.entrada_quantidade_recebida.get()
        self.saidamaterial = (self.entrada_saida_material.get())
        self.valorunitario = self.entrada_valor_unitario.get()
        self.valortotal = self.entrada_valor_total.get()


    def montarTabelas(self):
        self.conectar_banco()
        ### Criando tabela
        self.cursor.execute("""
            CREATE TABLE IF NOT EXISTS estoque(
                codigo INTEGER PRIMARY KEY,
                notafiscal CHAR(40) NOT NULL,
                codproduto CHAR(15),
                descricao CHAR(50),
                datavalidade DATE,
                unidade CHAR(20),
                quantidaderecebida DECIMAL,
                saidamaterial DECIMAL,
                quantidadetotal DECIMAL,
                valorunitario DECIMAL,
                valortotal DECIMAL
            );    
        """)
        self.conn.commit()
        self.desconectar_banco()

    def adicionar_cliente(self):
        self.variaveis()
        self.conectar_banco()

        # Converter os campos "Descrição" e "Unidade" para letras maiúsculas
        descricao = self.descricao.upper()
        unidade = self.unidade.upper()

        quantidadetotal = float(self.quantidaderecebida) - float(self.saidamaterial)

        # Inserir os dados na tabela "estoque"
        self.cursor.execute("""INSERT INTO estoque (notafiscal, codproduto, descricao, datavalidade, unidade, quantidaderecebida, saidamaterial, quantidadetotal, valorunitario, valortotal)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)""", (
            self.notafiscal, self.codproduto, descricao, self.data, unidade, self.quantidaderecebida,
            self.saidamaterial, quantidadetotal, self.valorunitario, self.valortotal))

        self.conn.commit()
        self.desconectar_banco()
        self.selecionar_na_lista()
        self.limpar_a_tela()

    def selecionar_na_lista(self):
        self.listaCli.delete(*self.listaCli.get_children())  ##limpar oque tem na lista
        self.conectar_banco()
        lista = self.cursor.execute("""SELECT codigo ,notafiscal,codproduto,descricao,datavalidade,unidade,quantidaderecebida,saidamaterial,valorunitario,valortotal,quantidadetotal FROM estoque
            ORDER BY notafiscal ASC; """)
        for i in lista:
            self.listaCli.insert("", END, values=i)
        self.desconectar_banco()

    def duplo_click(self, event):
        self.variaveis()
        self.limpar_a_tela()
        selected_item = self.listaCli.selection()

        if len(selected_item) == 1:
            col1, col2, col3, col4, col5, col6, col7, col8, col9, col10, col11 = self.listaCli.item(selected_item[0],
                                                                                                    'values')
            self.entrada_codigo.insert(END, col1)
            self.entrada_notafiscal.insert(END, col2)
            self.entrada_codproduto.insert(END, col3)
            self.entrada_descricao.insert(END, col4)
            self.entrada_data.insert(END, col5)
            self.entrada_unidade.insert(END, col6)
            self.entrada_quantidade_recebida.insert(END, col7)
            #self.entrada_saida_material.insert(END, col8)
            self.entrada_valor_unitario.insert(END, col9)
            self.entrada_valor_total.insert(END, col10)

    def apagar_cliente(self):
        self.variaveis()
        self.conectar_banco()
        self.cursor.execute("DELETE FROM estoque WHERE codigo = ?", (self.codigo,))
        self.conn.commit()
        self.desconectar_banco()
        self.limpar_a_tela()
        self.selecionar_na_lista()

    def alterar_cliente(self):
        self.variaveis()
        self.conectar_banco()

        # Converter os campos "Descrição" e "Unidade" para letras maiúsculas
        descricao = self.descricao.upper()
        unidade = self.unidade.upper()

        # Obter o valor anterior de saidamaterial antes da atualização
        self.cursor.execute("SELECT saidamaterial FROM estoque WHERE codigo = ?", (self.codigo,))
        resultado_anterior = self.cursor.fetchone()
        tempremovido_anterior = float(resultado_anterior[0]) if resultado_anterior else 0.0

        # Calcular tempremovido como a soma anterior e o novo valor de saidamaterial
        tempremovido = tempremovido_anterior + float(self.saidamaterial)

        # Atualizar os dados na tabela "estoque"
        self.cursor.execute("""
            UPDATE estoque
            SET notafiscal = ?, codproduto = ?, descricao = ?, datavalidade = ?,
                unidade = ?, quantidaderecebida = ?, saidamaterial = ?, quantidadetotal = ?,
                valorunitario = ?, valortotal = ?
            WHERE codigo = ?
        """, (
            self.notafiscal, self.codproduto, descricao, self.data, unidade, self.quantidaderecebida,
            tempremovido, float(self.quantidaderecebida) - tempremovido,
            self.valorunitario, self.valortotal, self.codigo
        ))

        self.conn.commit()
        self.desconectar_banco()
        self.selecionar_na_lista()
        self.limpar_a_tela()

    def buscar_cliente(self):
        self.variaveis()
        self.conectar_banco()
        self.listaCli.delete(*self.listaCli.get_children())
        descricao = self.entrada_descricao.get().upper()
        notafiscal = self.entrada_notafiscal.get().upper()
        codproduto = self.entrada_codproduto.get().upper()

        # Crie uma lista de condições para a pesquisa
        conditions = []
        params = []

        if descricao:
            conditions.append("INSTR(descricao, ?) > 0")
            params.append(descricao)

        if notafiscal:
            conditions.append("INSTR(notafiscal, ?) > 0")
            params.append(notafiscal)

        if codproduto:
            conditions.append("INSTR(codproduto, ?) > 0")
            params.append(codproduto)

        # Crie a consulta SQL com as condições
        if conditions:
            query = "SELECT codigo, notafiscal, codproduto, descricao, datavalidade, unidade, quantidaderecebida, saidamaterial, valorunitario, valortotal,quantidadetotal FROM estoque WHERE " + " AND ".join(
                conditions) + " ORDER BY descricao ASC"
            self.cursor.execute(query, params)

            buscarnomeCli = self.cursor.fetchall()
            for i in buscarnomeCli:
                self.listaCli.insert("", END, values=i)

        self.limpar_a_tela()
        self.desconectar_banco()


class aplicacoes(funcs):
    def __init__(self):
        self.janela = janela
        self.tela()
        self.frames_da_tela()
        self.botoes_de_cima()
        self.botoes_de_baixo()
        self.montarTabelas()
        self.selecionar_na_lista()
        janela.mainloop()

    def tela(self):
        self.janela.title("Coordenadoria de Saúde, Assistência Social e Religiosa")
        self.janela.configure(background="black")
        largura = janela.winfo_screenwidth()
        altura = janela.winfo_screenheight()
        janela.geometry(f"{largura}x{altura}+0+0")
        self.janela.resizable(True,True)
        self.janela.minsize(width=800, height=600)


    def frames_da_tela(self):
        self.frame_cima = Frame(self.janela, bd=4, bg='#dfe3ee', highlightbackground='#759fe6',
                                highlightthickness=3)
        self.frame_cima.place(relx=0.01, rely=0.01, relwidth=0.98, relheight=0.46)
        self.frame_baixo = Frame(self.janela, bd=4, bg='#dfe3ee', highlightbackground='#759fe6',
                                 highlightthickness=3)
        self.frame_baixo.place(relx=0.01, rely=0.48, relwidth=0.98, relheight=0.50)


    def botoes_de_cima(self):

        #botão de criar pdf
        self.pdf = Button(self.frame_cima, text="Exportar Excel", bd=2, bg="snow", fg="black",
                             font=("verdana", 8, 'bold'), command=self.exportar_excel)
        self.pdf.place(relx=0.858, rely=0.32, relheight=0.10, relwidth=0.07)

        ## Botão "Limpar"
        self.limpar = Button(self.frame_cima, text="Limpar", bd=2, bg="snow", fg="black",
                             font=("verdana", 8, 'bold'), command= self.limpar_a_tela)
        self.limpar.place(relx=0.93, rely=0.32, relheight=0.10, relwidth=0.07)
        ## Botão "Buscar"
        self.buscar = Button(self.frame_cima, text="Buscar", bd=2, bg="snow", fg="black",
                             font=("verdana", 8, 'bold'), command= self.buscar_cliente)
        self.buscar.place(relx=0.858, rely=0.21, relheight=0.1, relwidth=0.07)
        ## Botão "Alterar"
        self.alterar = Button(self.frame_cima, text="Alterar", bd=2, bg="snow", fg="black",
                              font=("verdana", 8, 'bold'),command=self.alterar_cliente)
        self.alterar.place(relx=0.93, rely=0.21, relheight=0.1, relwidth=0.07)
        ## Botão "Novo"
        self.novo = Button(self.frame_cima, text="Novo", bd=2, bg="snow", fg="black",
                           font=("verdana", 8, 'bold'), command = self.adicionar_cliente)
        self.novo.place(relx=0.858, rely=0.1, relheight=0.1, relwidth=0.07)
        ## Botão "Apagar"
        self.apagar = Button(self.frame_cima, text="Apagar", bd=2, bg="snow", fg="black",
                             font=("verdana", 8, 'bold'), command=self.apagar_cliente)
        self.apagar.place(relx=0.93, rely=0.1, relheight=0.1, relwidth=0.07)

        ##Label do notafiscal
        self.lb_notafiscal = Label(self.janela, text="Nota Fiscal:", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_notafiscal.place(relx=0.02, rely=0.03)

        ##Entry do notafiscal
        self.entrada_notafiscal = Entry(self.frame_cima)
        self.entrada_notafiscal.place(relx=0.006, rely=0.10, relwidth=0.10)

        ##Label do codproduto
        self.lb_codproduto = Label(self.janela, text="Código Produto:", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_codproduto.place(relx=0.02, rely=0.10)

        ##Entry do produto
        self.entrada_codproduto = Entry(self.frame_cima)
        self.entrada_codproduto.place(relx=0.006, rely=0.25, relwidth=0.10)

        ##Label do descrição
        self.lb_descricao = Label(self.janela, text="Descrição:", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_descricao.place(relx=0.02, rely=0.17)

        ##Entry do descrição
        self.entrada_descricao = Entry(self.frame_cima)
        self.entrada_descricao.place(relx=0.006, rely=0.41, relwidth=0.49)

        ##Label da data
        self.lb_data = Label(self.janela, text="Data Validade: dd/mm/aa", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_data.place(relx=0.15, rely=0.03)

        ##Entry da data
        self.entrada_data = DateEntry(self.frame_cima, borderwidth=2, date_pattern="dd.mm.yy")
        self.entrada_data.place(relx=0.14, rely=0.10, relwidth=0.10)

        ##Label da unidade
        self.lb_unidade = Label(self.janela, text="Unidade:", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_unidade.place(relx=0.15, rely=0.10)

        ##Entry da unidade
        self.entrada_unidade = Entry(self.frame_cima)
        self.entrada_unidade.place(relx=0.14, rely=0.25, relwidth=0.10)

        ##Label da quantidade recebida
        self.lb_quantidade_recebida = Label(self.janela, text="Quantidade recebida:", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_quantidade_recebida.place(relx=0.27, rely=0.03)

        ##Entry da quantidade recebida
        self.entrada_quantidade_recebida = Entry(self.frame_cima)
        self.entrada_quantidade_recebida.place(relx=0.265, rely=0.10, relwidth=0.10)

        ##Label da saida material
        self.lb_saida_material = Label(self.janela, text="Saída de material:", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_saida_material.place(relx=0.27, rely=0.10)

        ##Entry da saida material
        self.entrada_saida_material = Entry(self.frame_cima)
        self.entrada_saida_material.place(relx=0.265, rely=0.25, relwidth=0.10)

        ##Label da valor unitario
        self.lb_valor_unitario = Label(self.janela, text="Valor Unitário:", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_valor_unitario.place(relx=0.40, rely=0.03)

        ##Entry da valor unitario
        self.entrada_valor_unitario = Entry(self.frame_cima)
        self.entrada_valor_unitario.place(relx=0.398, rely=0.10, relwidth=0.10)

        ##Label Valor total
        self.lb_valor_total = Label(self.janela, text="Valor Total:", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_valor_total.place(relx=0.40, rely=0.10)

        ##Entry da valor total
        self.entrada_valor_total = Entry(self.frame_cima)
        self.entrada_valor_total.place(relx=0.398, rely=0.25, relwidth=0.10)


        #Label do código
        self.lb_codigo = Label(self.janela, text="Código:", bg='#dfe3ee', font=("Arial", 9, 'bold'))
        self.lb_codigo.place(relx=0.02, rely=0.33)

        ##Entry do valor
        self.entrada_codigo = Entry(self.frame_cima)
        self.entrada_codigo.place(relx=0.006, rely=0.80, relwidth=0.10)




    def botoes_de_baixo(self):
        self.listaCli = ttk.Treeview(self.frame_baixo, height=3, columns=("col1", "col2", "col3", "col4", "col5", "col6", "col7", "col8", "col9", "col10", "col11"))
        self.listaCli.heading("#0", text="")
        self.listaCli.heading("#1", text="Código")
        self.listaCli.heading("#2", text="Nota Fiscal")
        self.listaCli.heading("#3", text="Código Produto")
        self.listaCli.heading("#4", text="Descrição")
        self.listaCli.heading("#5", text="Data de Validade")
        self.listaCli.heading("#6", text="Unidade")
        self.listaCli.heading("#7", text="Quantidade Recebida")
        self.listaCli.heading("#8", text="Saída de Material")
        self.listaCli.heading("#9", text="Valor Unitário")
        self.listaCli.heading("#10", text="Valor Total")
        self.listaCli.heading("#11", text="Quantidade no Estoque")
        # Proporção de 500 para ordenar os itens
        self.listaCli.column("#0", width=1)
        self.listaCli.column("#1", width=1)
        self.listaCli.column("#2", width=100)
        self.listaCli.column("#3", width=100)
        self.listaCli.column("#4", width=300)
        self.listaCli.column("#5", width=100)
        self.listaCli.column("#6", width=100)
        self.listaCli.column("#7", width=100)
        self.listaCli.column("#8", width=100)
        self.listaCli.column("#9", width=100)
        self.listaCli.column("#10", width=100)
        self.listaCli.column("#11", width=100)

        self.listaCli.place(relx=0.001, rely=0.01, relwidth=0.97, relheight=0.99, )

        self.scrool = Scrollbar(self.janela, orient="vertical")
        self.listaCli.configure(yscroll=self.scrool.set)
        self.scrool.place(relx=0.96, rely=0.49, relwidth=0.02, relheight=0.48)

        self.listaCli.bind("<Double-1>", self.duplo_click)

aplicacoes()

