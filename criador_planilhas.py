import tkinter as tk
from tkinter import messagebox, ttk, filedialog
from openpyxl import Workbook
import pandas as pd  # Biblioteca para manipular arquivos de dados

# Função para adicionar dados à tabela
def adicionar_dados():
    nome = entry_nome.get()
    idade = entry_idade.get()
    cidade = entry_cidade.get()

    if not nome or not idade or not cidade:
        messagebox.showerror("Erro", "Todos os campos devem ser preenchidos!")
        return

    try:
        int(idade)  # Verifica se a idade é um número
    except ValueError:
        messagebox.showerror("Erro", "Idade deve ser um número!")
        return

    # Adiciona os dados na tabela
    tabela.insert("", "end", values=(nome, idade, cidade))

    # Limpa os campos de entrada
    entry_nome.delete(0, tk.END)
    entry_idade.delete(0, tk.END)
    entry_cidade.delete(0, tk.END)

# Função para adicionar dados em massa com pré-formatação
def adicionar_em_massa():
    dados = text_dados.get("1.0", tk.END).strip()
    if not dados:
        messagebox.showerror("Erro", "Nenhum dado inserido!")
        return

    linhas = dados.split("\n")
    for linha in linhas:
        # Tenta separar os dados por tabulação, vírgula ou espaço
        colunas = linha.replace(",", "\t").replace(" ", "\t").split("\t")
        colunas = [coluna.strip() for coluna in colunas if coluna.strip()]  # Remove espaços extras

        if len(colunas) != 3:
            messagebox.showerror("Erro", f"Formato inválido na linha: {linha}\nInsira Nome, Idade e Cidade.")
            return

        tabela.insert("", "end", values=(colunas[0], colunas[1], colunas[2]))

    # Limpa o campo de texto
    text_dados.delete("1.0", tk.END)
    messagebox.showinfo("Sucesso", "Dados adicionados com sucesso!")

# Função para criar a planilha com os dados da tabela
def criar_planilha():
    if not tabela.get_children():
        messagebox.showerror("Erro", "Nenhum dado para salvar!")
        return

    try:
        # Criando um arquivo Excel
        wb = Workbook()
        ws = wb.active
        ws.title = "Dados"

        # Adicionando cabeçalhos
        ws.append(["Nome", "Idade", "Cidade"])

        # Adicionando os dados da tabela
        for item in tabela.get_children():
            linha = tabela.item(item, "values")
            ws.append(linha)

        # Salvando o arquivo
        caminho_arquivo = "dados.xlsx"
        wb.save(caminho_arquivo)

        messagebox.showinfo("Sucesso", f"Planilha criada com sucesso: {caminho_arquivo}")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao criar a planilha: {e}")

# Função para importar dados de arquivos
def importar_dados():
    try:
        # Abrir o seletor de arquivos
        caminho_arquivo = filedialog.askopenfilename(
            title="Selecione um arquivo",
            filetypes=(("Arquivos CSV", "*.csv"), ("Arquivos Excel", "*.xlsx"), ("Arquivos de Texto", "*.txt"))
        )
        if not caminho_arquivo:
            return  # Se o usuário cancelar, não faz nada

        # Detecta o tipo de arquivo e lê os dados
        if caminho_arquivo.endswith(".csv"):
            dados = pd.read_csv(caminho_arquivo)
        elif caminho_arquivo.endswith(".xlsx"):
            dados = pd.read_excel(caminho_arquivo)
        elif caminho_arquivo.endswith(".txt"):
            dados = pd.read_csv(caminho_arquivo, delimiter="\t")
        else:
            messagebox.showerror("Erro", "Formato de arquivo não suportado!")
            return

        # Verifica se as colunas necessárias existem
        if not all(col in dados.columns for col in ["Nome", "Idade", "Cidade"]):
            messagebox.showerror("Erro", "O arquivo deve conter as colunas: Nome, Idade e Cidade.")
            return

        # Adiciona os dados à tabela
        for _, linha in dados.iterrows():
            tabela.insert("", "end", values=(linha["Nome"], linha["Idade"], linha["Cidade"]))

        messagebox.showinfo("Sucesso", "Dados importados com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao importar dados: {e}")

# Função para exibir o exemplo de entrada em massa
def exibir_exemplo():
    exemplo = "João\t25\tSão Paulo\nMaria\t30\tRio de Janeiro\nCarlos\t22\tBelo Horizonte"
    messagebox.showinfo("Exemplo de Entrada em Massa", f"Use o seguinte formato:\n\n{exemplo}")

# Função para tratar o fechamento da janela
def ao_fechar():
    if tabela.get_children():  # Verifica se há dados na tabela
        resposta = messagebox.askyesnocancel(
            "Salvar antes de sair",
            "Você deseja salvar os dados antes de sair?\n\nSim: Salvar e sair\nNão: Sair sem salvar\nCancelar: Voltar ao aplicativo"
        )
        if resposta is None:  # Cancelar
            return
        elif resposta:  # Sim
            criar_planilha()
    janela.destroy()  # Fecha a janela

# Criando a interface gráfica
janela = tk.Tk()
janela.title("Criador de Planilhas")
janela.geometry("700x600")
janela.configure(bg="#f0f0f0")  # Cor de fundo

# Título
titulo = tk.Label(janela, text="Criador de Planilhas", font=("Arial", 20, "bold"), bg="#f0f0f0", fg="#333")
titulo.pack(pady=10)

# Widgets para entrada de dados
frame_entrada = tk.Frame(janela, bg="#f0f0f0", bd=2, relief="groove")
frame_entrada.pack(pady=10, padx=10, fill="x")

label_nome = tk.Label(frame_entrada, text="Nome:", font=("Arial", 12), bg="#f0f0f0")
label_nome.grid(row=0, column=0, padx=5, pady=5, sticky="w")
entry_nome = tk.Entry(frame_entrada, font=("Arial", 12))
entry_nome.grid(row=0, column=1, padx=5, pady=5)

label_idade = tk.Label(frame_entrada, text="Idade:", font=("Arial", 12), bg="#f0f0f0")
label_idade.grid(row=1, column=0, padx=5, pady=5, sticky="w")
entry_idade = tk.Entry(frame_entrada, font=("Arial", 12))
entry_idade.grid(row=1, column=1, padx=5, pady=5)

label_cidade = tk.Label(frame_entrada, text="Cidade:", font=("Arial", 12), bg="#f0f0f0")
label_cidade.grid(row=2, column=0, padx=5, pady=5, sticky="w")
entry_cidade = tk.Entry(frame_entrada, font=("Arial", 12))
entry_cidade.grid(row=2, column=1, padx=5, pady=5)

botao_adicionar = tk.Button(frame_entrada, text="Adicionar", command=adicionar_dados, bg="#4caf50", fg="white", font=("Arial", 10))
botao_adicionar.grid(row=3, column=0, columnspan=2, pady=10)

# Campo de texto para entrada em massa
frame_massa = tk.Frame(janela, bg="#f0f0f0", bd=2, relief="groove")
frame_massa.pack(pady=10, padx=10, fill="x")

label_massa = tk.Label(frame_massa, text="Adicionar em Massa (Nome\tIdade\tCidade):", font=("Arial", 12), bg="#f0f0f0")
label_massa.pack()
text_dados = tk.Text(frame_massa, height=5, width=60, font=("Arial", 10))
text_dados.pack(pady=5)

# Botão para exibir o exemplo
botao_exemplo = tk.Button(frame_massa, text="Exibir Exemplo", command=exibir_exemplo, bg="#9c27b0", fg="white", font=("Arial", 10))
botao_exemplo.pack(pady=5)

botao_massa = tk.Button(frame_massa, text="Adicionar em Massa", command=adicionar_em_massa, bg="#ff9800", fg="white", font=("Arial", 10))
botao_massa.pack(pady=5)

# Botão para importar dados
botao_importar = tk.Button(frame_massa, text="Importar Arquivo", command=importar_dados, bg="#3f51b5", fg="white", font=("Arial", 10))
botao_importar.pack(pady=5)

# Tabela para exibir os dados
frame_tabela = tk.Frame(janela, bg="#f0f0f0", bd=2, relief="groove")
frame_tabela.pack(pady=10, padx=10, fill="both", expand=True)

colunas = ("Nome", "Idade", "Cidade")
tabela = ttk.Treeview(frame_tabela, columns=colunas, show="headings", height=10)
tabela.pack(fill="both", expand=True)

for coluna in colunas:
    tabela.heading(coluna, text=coluna)
    tabela.column(coluna, width=200)

# Botão para criar a planilha
botao_criar = tk.Button(janela, text="Criar Planilha", command=criar_planilha, bg="#2196f3", fg="white", font=("Arial", 12))
botao_criar.pack(pady=10)

# Configura o evento de fechamento da janela
janela.protocol("WM_DELETE_WINDOW", ao_fechar)

# Iniciando o loop da interface gráfica
janela.mainloop()