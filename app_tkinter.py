import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import threading
import re
import time

# Função para carregar a planilha
def carregar_planilha():
    caminho_arquivo = filedialog.askopenfilename(filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho_arquivo:
        label_status.config(text=f"Planilha carregada: {caminho_arquivo}")
        try:
            global df
            df = pd.read_excel(caminho_arquivo)  # Carrega a planilha
            label_status.config(text="Planilha carregada com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")
    else:
        label_status.config(text="Nenhuma planilha foi carregada.")

# Função para usar o Selenium para buscar preços no Google
def buscar_precos_google(driver, titulo, autor):
    # Criar a URL de pesquisa no Google
    query = f"{titulo} {autor} preço livro"
    url = f"https://www.google.com/search?q={query.replace(' ', '+')}"

    # Abrir a página de busca no Google (reutilizando o driver)
    driver.get(url)

    try:
        # Aguardar até que os preços sejam carregados
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, "//span[contains(text(), 'R$')]")))

        # Usar regex para encontrar todos os preços na página
        precos_texto = driver.page_source
        precos = re.findall(r'R\$ ?(\d+,\d+)', precos_texto)
        precos = [float(preco.replace(',', '.')) for preco in precos]
        
        if precos:
            return min(precos)  # Retorna o menor preço encontrado
    except Exception as e:
        print(f"Erro ao buscar preços para {titulo}: {e}")
    
    return None  # Se nenhum preço for encontrado

# Função para buscar os preços no Google e salvar na planilha
def buscar_precos():
    if df is None:
        messagebox.showwarning("Aviso", "Carregue uma planilha primeiro.")
        return

    # Configurar o ChromeDriver (assumindo que você já baixou o ChromeDriver e ele está no PATH)
    driver = webdriver.Chrome()

    # Garantir que a coluna 'Pesquisa' seja do tipo 'object' para aceitar strings
    if 'Pesquisa' in df.columns:
        df['Pesquisa'] = df['Pesquisa'].astype(object)

    label_status.config(text="Buscando preços no Google...")
    progresso['value'] = 0
    progresso['maximum'] = len(df)

    for index, row in df.iterrows():
        titulo = row['Título']
        autor = row['Autor']
        
        # Realizar a busca de preços no Google reutilizando o navegador
        preco = buscar_precos_google(driver, titulo, autor)
        
        # Atualiza o DataFrame com o menor preço encontrado ou uma mensagem de erro
        if preco:
            df.at[index, 'Pesquisa'] = f"R$ {preco:.2f}"
        else:
            df.at[index, 'Pesquisa'] = "Preço não encontrado"
        
        # Atualizar a barra de progresso
        progresso['value'] = index + 1

    driver.quit()  # Fechar o navegador ao final do processo

    # Salva a planilha atualizada com os preços
    caminho_salvar = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Arquivos Excel", "*.xlsx")])
    if caminho_salvar:
        df.to_excel(caminho_salvar, index=False)
        messagebox.showinfo("Info", f"Planilha salva com sucesso em: {caminho_salvar}")
    else:
        messagebox.showinfo("Info", "Busca de preços concluída, mas a planilha não foi salva.")

    label_status.config(text="Busca de preços concluída!")

# Função para iniciar a busca em segundo plano
def iniciar_busca():
    threading.Thread(target=buscar_precos).start()

# Interface gráfica com Tkinter
janela = tk.Tk()
janela.title("Pesquisa de Preços de Livros")

# Botão para carregar a planilha
btn_carregar = tk.Button(janela, text="Carregar Planilha", command=carregar_planilha)
btn_carregar.pack(pady=10)

# Barra de progresso
progresso = ttk.Progressbar(janela, orient="horizontal", length=300, mode="determinate")
progresso.pack(pady=10)

# Botão para iniciar a busca de preços
btn_buscar = tk.Button(janela, text="Buscar Preços", command=iniciar_busca)
btn_buscar.pack(pady=10)

# Label de status
label_status = tk.Label(janela, text="Nenhuma planilha carregada.")
label_status.pack(pady=10)

# Iniciar o loop principal da interface
janela.mainloop()
