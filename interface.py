import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import pyperclip  # pip install pyperclip


def limpar_valor(valor):
    """
    Normaliza valores vindos do Excel:
    - Se for string: remove espaços extras nas pontas e substitui strings só de espaço por ''.
    - Se for float inteiro (ex: 123.0): vira int (123).
    """
    if isinstance(valor, str):
        v = valor.strip()
        return v if v != "" else ""
    if isinstance(valor, float) and valor.is_integer():
        return int(valor)
    return valor


def criar_interface(caminho_arquivo_dados):
    """
    Cria a interface gráfica com um combobox para pesquisa de clientes e exibição de dados.
    """
    try:
        # Lê o Excel
        df_original = pd.read_excel(caminho_arquivo_dados, sheet_name='DADOS')

        # Substitui NaN por '' antes de limpar
        df_original = df_original.fillna('')

        # LIMPA todo o DataFrame (cada célula)
        df_original = df_original.applymap(limpar_valor)

        # Debug opcional
        print("Colunas disponíveis na aba 'DADOS':", df_original.columns.tolist())

        # Garante que existe 'Nome'
        if 'Nome' not in df_original.columns:
            messagebox.showerror(
                "Erro de Coluna",
                "A coluna 'Nome' não foi encontrada na aba 'DADOS' do arquivo Excel. "
                "Verifique o nome exato da coluna ou ajuste o código."
            )
            return

    except FileNotFoundError:
        messagebox.showerror(
            "Erro",
            f"Arquivo '{caminho_arquivo_dados}' não encontrado. "
            f"Certifique-se de que o arquivo está na mesma pasta do script."
        )
        return
    except KeyError:
        messagebox.showerror(
            "Erro de Aba",
            f"A aba 'DADOS' não foi encontrada no arquivo '{caminho_arquivo_dados}'. "
            f"Por favor, verifique se o nome da aba está correto."
        )
        return
    except Exception as e:
        messagebox.showerror("Erro", f"Ocorreu um erro ao ler o arquivo: {e}")
        return

    # ==== INÍCIO DA INTERFACE ====
    root = tk.Tk()
    root.title("Copiador de Dados de Clientes")
    root.geometry("900x250")

    frame_principal = ttk.Frame(root)
    frame_principal.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    # ----------- SEÇÃO ESQUERDA (PESQUISA) -----------
    frame_pesquisa = ttk.LabelFrame(frame_principal, text="Pesquisar Cliente")
    frame_pesquisa.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5, pady=5)

    lbl_pesquisa = ttk.Label(frame_pesquisa, text="Selecione ou digite o nome:")
    lbl_pesquisa.pack(pady=5)

    combobox_clientes = ttk.Combobox(frame_pesquisa, width=50, state="readonly")
    combobox_clientes.pack(pady=5)
    combobox_clientes.focus_set()

    # Monta a lista de nomes já LIMPOS (sem espaços)
    nomes_clientes_todos = (
        df_original['Nome']
        .astype(str)
        .map(lambda x: x.strip())
        .map(lambda x: x if x != "" else "(SEM NOME)")
        .tolist()
    )

    # Para evitar nomes duplicados que só diferem por espaço, vamos normalizar e ordenar únicos
    nomes_clientes_todos = sorted(dict.fromkeys(nomes_clientes_todos))

    combobox_clientes['values'] = nomes_clientes_todos

    # ----------- SEÇÃO DIREITA (DADOS DO CLIENTE) -----------
    frame_exibicao_dados = ttk.LabelFrame(frame_principal, text="Dados do Cliente Selecionado")
    frame_exibicao_dados.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=5, pady=5)

    variaveis_dados = {}

    # Ajuste esta lista para corresponder aos nomes EXATOS das colunas no Excel
    campos_para_exibir = ["Nome", "CPF", "Inscrição Estadual", "SENHA IMA", "TELEFONE", "EMAIL"]

    for i, campo in enumerate(campos_para_exibir):
        ttk.Label(frame_exibicao_dados, text=f"{campo}:").grid(
            row=i, column=0, sticky="w", padx=5, pady=2
        )

        var = tk.StringVar()
        entry = ttk.Entry(
            frame_exibicao_dados,
            textvariable=var,
            width=50,
            state='readonly'
        )
        entry.grid(row=i, column=1, sticky="ew", padx=5, pady=2)
        variaveis_dados[campo] = var

        botao_copiar = ttk.Button(
            frame_exibicao_dados,
            text="Copiar",
            command=lambda v=var: copiar_para_area_transferencia(v.get())
        )
        botao_copiar.grid(row=i, column=2, padx=5, pady=2)

    def copiar_para_area_transferencia(texto):
        """Copia o texto fornecido para a área de transferência."""
        try:
            texto_limpo = str(texto).strip()
            pyperclip.copy(texto_limpo)
        except pyperclip.PyperclipException as e:
            messagebox.showerror(
                "Erro ao Copiar",
                f"Não foi possível copiar o texto: {e}\n"
                f"Certifique-se de ter um copiador de área de transferência instalado (por exemplo, xclip ou xsel no Linux)."
            )

    def normalizar_ie(valor_ie):
        """
        Trata Inscrição Estadual:
        - Remove '.0'
        - Preenche com zeros à esquerda até 13 dígitos
        - Remove espaços excedentes
        """
        valor_ie = str(valor_ie).replace('.0', '').strip()
        if valor_ie == "":
            return ""
        return valor_ie.zfill(13)

    def mostrar_dados_cliente_selecionado(event):
        """Exibe os dados do cliente selecionado no combobox."""
        nome_selecionado = combobox_clientes.get().strip()

        # Limpa visualmente todos os campos antes de atualizar
        for var in variaveis_dados.values():
            var.set("")

        if not nome_selecionado:
            return

        # Caso "(SEM NOME)" esteja no combobox, não vamos tentar bater com Nome
        if nome_selecionado == "(SEM NOME)":
            cliente_mask = df_original['Nome'].astype(str).map(lambda x: x.strip() == "")
        else:
            cliente_mask = df_original['Nome'].astype(str).map(lambda x: x.strip()) == nome_selecionado

        cliente_data = df_original[cliente_mask]

        if cliente_data.empty:
            # Nada encontrado → mostra mensagem clara
            for campo in campos_para_exibir:
                variaveis_dados[campo].set("Cliente não encontrado")
            return

        # Usa a primeira linha correspondente
        cliente_data = cliente_data.iloc[0]

        for campo in campos_para_exibir:
            # pega o valor bruto
            valor = cliente_data.get(campo, "")

            # passa pelo mesmo limpador geral
            valor = limpar_valor(valor)

            # tratamento especial pra IE
            if campo == "Inscrição Estadual":
                valor = normalizar_ie(valor)

            # garante que vai como string, sem espaço visual sobrando
            valor_final = str(valor).strip()

            # evita exibir "nan", "None", etc.
            if valor_final.lower() in ["nan", "none"]:
                valor_final = ""

            variaveis_dados[campo].set(valor_final)

    # Quando o usuário seleciona um cliente na lista
    combobox_clientes.bind("<<ComboboxSelected>>", mostrar_dados_cliente_selecionado)

    def on_combobox_keyrelease(event):
        """
        Quando o usuário começa a digitar no combobox (mesmo ele sendo readonly, você pode mudar isso depois se quiser),
        se o texto não é exatamente um nome conhecido, limpamos os campos à direita.
        """
        texto_digitado = combobox_clientes.get().strip().lower()
        lista_normalizada = [n.lower().strip() for n in nomes_clientes_todos]
        if texto_digitado not in lista_normalizada:
            for var in variaveis_dados.values():
                var.set("")

    combobox_clientes.bind("<KeyRelease>", on_combobox_keyrelease)

    root.mainloop()


if __name__ == "__main__":
    criar_interface(r"CONFERENCIA LIVRO CAIXA 2025.xlsx")
