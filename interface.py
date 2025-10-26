import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import pyperclip # Você ainda precisará desta biblioteca: pip install pyperclip

def criar_interface(caminho_arquivo_dados):
    """
    Cria a interface gráfica com um combobox para pesquisa de clientes e exibição de dados.

    Args:
        caminho_arquivo_dados (str): O caminho para o arquivo Excel (.xlsx) com os dados.
    """
    try:
        # Lendo o arquivo Excel, especificando a aba 'DADOS'
        df_original = pd.read_excel(caminho_arquivo_dados, sheet_name='DADOS')
        # Preenchendo valores NaN com string vazia para evitar erros de tipo na pesquisa e exibição
        df_original = df_original.fillna('')

        # --- LINHA PARA AJUDAR NA DEPURACAO (DEBUG) ---
        print("Colunas disponíveis na aba 'DADOS':", df_original.columns.tolist())
        # ---------------------------------------------

        # Assegura que a coluna 'Nome' existe. Se o nome da coluna de nomes de clientes
        # for diferente no seu Excel (ex: 'Nome Completo', 'Cliente'), ajuste 'Nome' abaixo
        # para o nome EXATO da coluna no seu arquivo.
        if 'Nome' not in df_original.columns:
            messagebox.showerror("Erro de Coluna", "A coluna 'Nome' não foi encontrada na aba 'DADOS' do arquivo Excel. Verifique o nome exato da coluna ou ajuste o código.")
            return

    except FileNotFoundError:
        messagebox.showerror("Erro", f"Arquivo '{caminho_arquivo_dados}' não encontrado. Certifique-se de que o arquivo está na mesma pasta do script.")
        return
    except KeyError:
        messagebox.showerror("Erro de Aba", f"A aba 'DADOS' não foi encontrada no arquivo '{caminho_arquivo_dados}'. Por favor, verifique se o nome da aba está correto.")
        return
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao ler o arquivo: {e}\nDetalhes: {e}")
        return

    root = tk.Tk()
    root.title("Copiador de Dados de Clientes")
    root.geometry("900x250") # Mantém a dimensão que você havia definido

    # --- Frame principal dividido em duas seções (esquerda para pesquisa, direita para dados) ---
    frame_principal = ttk.Frame(root)
    frame_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # --- Seção Esquerda: Pesquisa com Combobox ---
    frame_pesquisa = ttk.LabelFrame(frame_principal, text="Pesquisar Cliente")
    frame_pesquisa.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

    lbl_pesquisa = ttk.Label(frame_pesquisa, text="Selecione ou digite o nome:")
    lbl_pesquisa.pack(pady=5)
    
    # NOVO: Combobox para pesquisa e seleção de clientes
    combobox_clientes = ttk.Combobox(frame_pesquisa, width=50, state="readonly") 
    combobox_clientes.pack(pady=5)
    combobox_clientes.focus_set() # Foca no combobox ao iniciar

    # Preenche o combobox com todos os nomes de clientes inicialmente
    nomes_clientes_todos = sorted(df_original['Nome'].astype(str).tolist())
    combobox_clientes['values'] = nomes_clientes_todos

    # --- Seção Direita: Exibição e Botões de Copiar ---
    frame_exibicao_dados = ttk.LabelFrame(frame_principal, text="Dados do Cliente Selecionado")
    frame_exibicao_dados.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)

    variaveis_dados = {}
    # Nomes das colunas que você quer exibir e copiar.
    # AJUSTE ESTA LISTA COM OS NOMES EXATOS DAS SUAS COLUNAS NO SEU ARQUIVO EXCEL.
    campos_para_exibir = ["Nome", "CPF", "Inscrição Estadual", "SENHA IMA", "TELEFONE", "EMAIL"]

    for i, campo in enumerate(campos_para_exibir):
        ttk.Label(frame_exibicao_dados, text=f"{campo}:").grid(row=i, column=0, sticky="w", padx=5, pady=2)
        var = tk.StringVar()
        entry = ttk.Entry(frame_exibicao_dados, textvariable=var, width=50, state='readonly')
        entry.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
        variaveis_dados[campo] = var

        botao_copiar = ttk.Button(frame_exibicao_dados, text="Copiar", command=lambda v=var: copiar_para_area_transferencia(v.get()))
        botao_copiar.grid(row=i, column=2, padx=5, pady=2)

    def copiar_para_area_transferencia(texto):
        """Copia o texto fornecido para a área de transferência sem exibir mensagem."""
        try:
            pyperclip.copy(str(texto))
        except pyperclip.PyperclipException as e:
            messagebox.showerror("Erro ao Copiar", f"Não foi possível copiar o texto: {e}\nCertifique-se de ter um copiador de área de transferência instalado (por exemplo, xclip ou xsel no Linux).")

    def mostrar_dados_cliente_selecionado(event):
        """Exibe os dados do cliente selecionado no combobox."""
        nome_selecionado = combobox_clientes.get()
        
        # Limpa os campos de exibição de dados
        for var in variaveis_dados.values():
            var.set("")

        if nome_selecionado:
            # Encontra a linha correspondente no DataFrame original
            # Certifica-se de que a comparação de nomes lida com tipos e espaços
            cliente_data = df_original[df_original['Nome'].astype(str).str.strip() == str(nome_selecionado).strip()]
            if not cliente_data.empty:
                cliente_data = cliente_data.iloc[0]
                for campo in campos_para_exibir:
                    valor = cliente_data.get(campo, "N/A")
                    
                    # --- ALTERAÇÃO 1: Formatar números sem '.0' ---
                    if isinstance(valor, float) and valor.is_integer():
                        valor = int(valor) # Converte float (ex: 123.0) para int (ex: 123)
                    
                    # --- ALTERAÇÃO 2: Formatar Inscrição Estadual com zeros à esquerda ---
                    if campo == "Inscrição Estadual":
                        # Converte para string e remove o ".0" se existir (caso ainda venha como float)
                        valor_str = str(valor).replace('.0', '') 
                        # Define o comprimento desejado (ex: 13 para "0011793890196")
                        # Se suas IEs tiverem comprimentos variáveis e o "00" for sempre prefixo,
                        # você pode precisar de uma lógica mais complexa.
                        # Assumindo que o "00" inicial é para totalizar um certo número de dígitos (ex: 13)
                        # ou que sempre deve ter pelo menos 2 zeros à esquerda se não tiver.
                        
                        # Exemplo de formatação para 13 dígitos totais, preenchendo com zeros à esquerda
                        # Se o número original for 11793890196 (11 dígitos), ele se tornará 0011793890196
                        # Se o número original já tiver 13 dígitos ou mais, não será alterado.
                        valor = valor_str.zfill(13) 
                        
                    # ---------------------------------------------------------------------

                    variaveis_dados[campo].set(valor)
            else:
                for campo in campos_para_exibir:
                    variaveis_dados[campo].set("Cliente não encontrado")


    # Vincula o evento de seleção do combobox à função para mostrar os dados
    combobox_clientes.bind("<<ComboboxSelected>>", mostrar_dados_cliente_selecionado)
    
    # Opcional: Para limpar os campos de exibição ao digitar no combobox e o nome não estiver completo
    def on_combobox_keyrelease(event):
        texto_digitado = combobox_clientes.get().strip().lower()
        # Se o texto digitado não corresponder a nenhum nome exato na lista, limpa os campos
        if texto_digitado not in [n.lower() for n in nomes_clientes_todos]:
             for var in variaveis_dados.values():
                var.set("")

    combobox_clientes.bind("<KeyRelease>", on_combobox_keyrelease)


    root.mainloop()

if __name__ == "__main__":
    # O caminho completo para o arquivo Excel.
    criar_interface(r"CONFERENCIA LIVRO CAIXA 2025.xlsx")