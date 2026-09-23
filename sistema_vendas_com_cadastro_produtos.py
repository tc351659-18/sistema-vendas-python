import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill

ARQUIVO_DADOS = "dados_sistema.xlsx"
ARQUIVO_RELATORIO = "relatorio_vendas.xlsx"

vendas = []
produtos = {}


def moeda(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def atualizar_dashboard():
    faturamento = sum(v["faturamento"] for v in vendas)
    custos = sum(v["custo_total"] for v in vendas)
    lucro = sum(v["lucro"] for v in vendas)
    quantidade = sum(v["quantidade"] for v in vendas)

    lbl_faturamento.config(text=moeda(faturamento))
    lbl_custos.config(text=moeda(custos))
    lbl_lucro.config(text=moeda(lucro))
    lbl_itens.config(text=str(quantidade))
    lbl_vendas.config(text=str(len(vendas)))
    lbl_produtos.config(text=str(len(produtos)))


def atualizar_lista_produtos():
    for item in tree_produtos.get_children():
        tree_produtos.delete(item)

    for nome, dados in sorted(produtos.items()):
        estoque = dados["estoque"]
        minimo = dados["estoque_minimo"]

        if estoque <= 0:
            status = "ESGOTADO"
        elif estoque <= minimo:
            status = "ESTOQUE BAIXO"
        else:
            status = "OK"

        tree_produtos.insert(
            "", "end",
            values=(
                nome,
                moeda(dados["preco"]),
                moeda(dados["custo"]),
                estoque,
                minimo,
                status
            )
        )

    atualizar_dashboard()
    atualizar_combo_produtos()


def atualizar_combo_produtos():
    nomes = sorted(produtos.keys())
    combo_produto["values"] = nomes

    if combo_produto.get() not in nomes:
        combo_produto.set("")

    atualizar_dados_venda()


def atualizar_dados_venda(event=None):
    nome = combo_produto.get().strip()

    if nome in produtos:
        dados = produtos[nome]

        entry_preco_venda.delete(0, tk.END)
        entry_preco_venda.insert(0, str(dados["preco"]))

        entry_custo_venda.delete(0, tk.END)
        entry_custo_venda.insert(0, str(dados["custo"]))

        lbl_estoque_disponivel.config(
            text=f"Estoque disponível: {dados['estoque']}"
        )
    else:
        entry_preco_venda.delete(0, tk.END)
        entry_custo_venda.delete(0, tk.END)
        lbl_estoque_disponivel.config(text="Estoque disponível: -")


def atualizar_estoque():
    for item in tree_estoque.get_children():
        tree_estoque.delete(item)

    vendidos_por_produto = {}

    for v in vendas:
        vendidos_por_produto[v["produto"]] = (
            vendidos_por_produto.get(v["produto"], 0) + v["quantidade"]
        )

    for nome, dados in sorted(produtos.items()):
        vendido = vendidos_por_produto.get(nome, 0)
        atual = dados["estoque"]
        minimo = dados["estoque_minimo"]

        if atual <= 0:
            status = "ESGOTADO"
        elif atual <= minimo:
            status = "ESTOQUE BAIXO"
        else:
            status = "OK"

        tree_estoque.insert(
            "", "end",
            values=(nome, dados["estoque_inicial"], vendido, atual, minimo, status)
        )


def atualizar_historico():
    for item in tree_historico.get_children():
        tree_historico.delete(item)

    busca = entry_busca.get().strip().lower()

    for i in reversed(range(len(vendas))):
        v = vendas[i]

        if busca and busca not in v["produto"].lower():
            continue

        tree_historico.insert(
            "",
            "end",
            iid=str(i),
            values=(
    v.get("data", "Sem data"),
    v["produto"],
    moeda(v["preco"]),
    moeda(v["custo"]),
    v["quantidade"],
    moeda(v["faturamento"]),
    moeda(v["lucro"])
)
        )


def salvar_dados():
    wb = Workbook()

    ws = wb.active
    ws.title = "Produtos"
    ws.append([
        "Produto",
        "Preço de Venda",
        "Custo",
        "Estoque Inicial",
        "Estoque Atual",
        "Estoque Mínimo"
    ])

    for nome, dados in sorted(produtos.items()):
        ws.append([
            nome,
            dados["preco"],
            dados["custo"],
            dados["estoque_inicial"],
            dados["estoque"],
            dados["estoque_minimo"]
        ])

    ws2 = wb.create_sheet("Vendas")
    ws2.append([
    "Data/Hora",
    "Produto",
    "Preço",
    "Custo",
    "Quantidade",
    "Faturamento",
    "Custo Total",
    "Lucro"
])

    for v in vendas: 
        ws2.append([
        v.get("data", ""),
        v["produto"],
        v["preco"],
        v["custo"],
        v["quantidade"],
        v["faturamento"],
        v["custo_total"],
        v["lucro"]
    ])

    for sheet in wb.worksheets:
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="center")

        for col in sheet.columns:
            letra = col[0].column_letter
            maior = max(len(str(c.value or "")) for c in col)
            sheet.column_dimensions[letra].width = min(maior + 3, 30)

    wb.save(ARQUIVO_DADOS)


def carregar_dados():
    global produtos, vendas

    try:
        wb = load_workbook(ARQUIVO_DADOS, data_only=True)

        produtos = {}
        vendas = []

        if "Produtos" in wb.sheetnames:
            ws = wb["Produtos"]

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue

                nome = row[0]
                preco = row[1] if len(row) > 1 else 0
                custo = row[2] if len(row) > 2 else 0
                estoque_inicial = row[3] if len(row) > 3 else 0
                estoque_atual = row[4] if len(row) > 4 else estoque_inicial
                estoque_minimo = row[5] if len(row) > 5 else 0

                produtos[str(nome)] = {
                    "preco": float(preco or 0),
                    "custo": float(custo or 0),
                    "estoque_inicial": int(estoque_inicial or 0),
                    "estoque": int(estoque_atual or 0),
                    "estoque_minimo": int(estoque_minimo or 0)
                }

        elif "Estoque" in wb.sheetnames:
            ws = wb["Estoque"]

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue

                nome = row[0]
                estoque_inicial = row[1] if len(row) > 1 else 0
                estoque_atual = row[2] if len(row) > 2 else estoque_inicial

                produtos[str(nome)] = {
                    "preco": 0,
                    "custo": 0,
                    "estoque_inicial": int(estoque_inicial or 0),
                    "estoque": int(estoque_atual or 0),
                    "estoque_minimo": 0
                }

        if "Vendas" in wb.sheetnames:
            ws = wb["Vendas"]

            for row in ws.iter_rows(min_row=2, values_only=True):
                if not row[0]:
                    continue

                if len(row) >= 8:
                    data, nome, preco, custo, quantidade, faturamento, custo_total, lucro = row

                    vendas.append({
                        "data": str(data or ""),
                        "produto": str(nome),
                        "preco": float(preco or 0),
                        "custo": float(custo or 0),
                        "quantidade": int(quantidade or 0),
                        "faturamento": float(faturamento or 0),
                        "custo_total": float(custo_total or 0),
                        "lucro": float(lucro or 0)
                    })

                elif len(row) >= 7:
                    nome, preco, custo, quantidade, faturamento, custo_total, lucro = row

                    vendas.append({
                        "data": "",
                        "produto": str(nome),
                        "preco": float(preco or 0),
                        "custo": float(custo or 0),
                        "quantidade": int(quantidade or 0),
                        "faturamento": float(faturamento or 0),
                        "custo_total": float(custo_total or 0),
                        "lucro": float(lucro or 0)
                    })

    except FileNotFoundError:
        pass

    except Exception as e:
        messagebox.showwarning(
            "Aviso",
            f"Não foi possível carregar os dados:\n{e}"
        )

    atualizar_dashboard()
    atualizar_lista_produtos()
    atualizar_estoque()
    atualizar_historico()

def cadastrar_produto():
    nome = entry_nome_produto.get().strip()

    try:
        preco = float(entry_preco_produto.get().replace(",", "."))
        custo = float(entry_custo_produto.get().replace(",", "."))
        estoque_inicial = int(entry_estoque_produto.get())
        estoque_minimo = int(entry_minimo_produto.get())
    except ValueError:
        messagebox.showerror(
            "Erro",
            "Preencha preço, custo, estoque e estoque mínimo com valores válidos."
        )
        return

    if not nome:
        messagebox.showerror("Erro", "Digite o nome do produto.")
        return

    if preco < 0 or custo < 0:
        messagebox.showerror("Erro", "Preço e custo não podem ser negativos.")
        return

    if estoque_inicial < 0 or estoque_minimo < 0:
        messagebox.showerror("Erro", "Estoque não pode ser negativo.")
        return

    if nome in produtos:
        resposta = messagebox.askyesno(
            "Produto já cadastrado",
            f"O produto '{nome}' já existe.\n\n"
            "Deseja atualizar os dados dele?"
        )

        if not resposta:
            return

        # Mantém o estoque atual quando apenas preço/custo/mínimo são alterados
        estoque_atual = produtos[nome]["estoque"]
        estoque_original = produtos[nome]["estoque_inicial"]
    else:
        estoque_atual = estoque_inicial
        estoque_original = estoque_inicial

    produtos[nome] = {
        "preco": preco,
        "custo": custo,
        "estoque_inicial": estoque_original,
        "estoque": estoque_atual,
        "estoque_minimo": estoque_minimo
    }

    salvar_dados()
    atualizar_lista_produtos()
    atualizar_estoque()
    limpar_formulario_produto()

    messagebox.showinfo(
        "Produto salvo",
        f"Produto '{nome}' cadastrado com sucesso!"
    )


def selecionar_produto(event=None):
    selecionado = tree_produtos.selection()

    if not selecionado:
        return

    valores = tree_produtos.item(selecionado[0], "values")
    nome = valores[0]

    if nome not in produtos:
        return

    dados = produtos[nome]

    entry_nome_produto.delete(0, tk.END)
    entry_nome_produto.insert(0, nome)

    entry_preco_produto.delete(0, tk.END)
    entry_preco_produto.insert(0, str(dados["preco"]))

    entry_custo_produto.delete(0, tk.END)
    entry_custo_produto.insert(0, str(dados["custo"]))

    entry_estoque_produto.delete(0, tk.END)
    entry_estoque_produto.insert(0, str(dados["estoque"]))

    entry_minimo_produto.delete(0, tk.END)
    entry_minimo_produto.insert(0, str(dados["estoque_minimo"]))


def limpar_formulario_produto():
    entry_nome_produto.delete(0, tk.END)
    entry_preco_produto.delete(0, tk.END)
    entry_custo_produto.delete(0, tk.END)
    entry_estoque_produto.delete(0, tk.END)
    entry_minimo_produto.delete(0, tk.END)
    entry_estoque_produto.insert(0, "0")
    entry_minimo_produto.insert(0, "5")
    entry_nome_produto.focus()


def registrar_venda():
    nome = combo_produto.get().strip()

    if nome not in produtos:
        messagebox.showerror(
            "Erro",
            "Selecione um produto cadastrado."
        )
        return

    try:
        quantidade = int(entry_quantidade_venda.get())
    except ValueError:
        messagebox.showerror(
            "Erro",
            "Digite uma quantidade válida."
        )
        return

    if quantidade <= 0:
        messagebox.showerror(
            "Erro",
            "A quantidade deve ser maior que zero."
        )
        return

    dados = produtos[nome]

    if quantidade > dados["estoque"]:
        messagebox.showerror(
            "Estoque insuficiente",
            f"Produto: {nome}\n"
            f"Estoque disponível: {dados['estoque']}\n"
            f"Quantidade solicitada: {quantidade}"
        )
        return

    preco = dados["preco"]
    custo = dados["custo"]

    faturamento = preco * quantidade
    custo_total = custo * quantidade
    lucro = faturamento - custo_total

    dados["estoque"] -= quantidade

    vendas.append({
    "data": datetime.now().strftime("%d/%m/%Y %H:%M:%S"),
    "produto": nome,
    "preco": preco,
    "custo": custo,
    "quantidade": quantidade,
    "faturamento": faturamento,
    "custo_total": custo_total,
    "lucro": lucro
})

    salvar_dados()
    atualizar_dashboard()
    atualizar_lista_produtos()
    atualizar_estoque()
    atualizar_historico()
    atualizar_dados_venda()

    entry_quantidade_venda.delete(0, tk.END)
    entry_quantidade_venda.insert(0, "1")

    messagebox.showinfo(
        "Venda registrada",
        f"Venda registrada com sucesso!\n\n"
        f"Produto: {nome}\n"
        f"Quantidade: {quantidade}\n"
        f"Faturamento: {moeda(faturamento)}\n"
        f"Lucro: {moeda(lucro)}\n"
        f"Estoque restante: {dados['estoque']}"
    )

def excluir_venda():
    selecionado = tree_historico.selection()

    if not selecionado:
        messagebox.showwarning(
            "Excluir venda",
            "Selecione uma venda no histórico primeiro."
        )
        return

    item = selecionado[0]

    try:
        indice = int(item)
        venda = vendas[indice]
    except (ValueError, IndexError):
        messagebox.showerror(
            "Erro",
            "Não foi possível identificar a venda selecionada."
        )
        return

    confirmar = messagebox.askyesno(
        "Confirmar exclusão",
        f"Tem certeza que deseja excluir esta venda?\n\n"
        f"Produto: {venda['produto']}\n"
        f"Quantidade: {venda['quantidade']}\n"
        f"Faturamento: {moeda(venda['faturamento'])}\n\n"
        f"A quantidade será devolvida ao estoque."
    )

    if not confirmar:
        return

    produto = venda["produto"]

    if produto in produtos:
        produtos[produto]["estoque"] += venda["quantidade"]

    vendas.pop(indice)

    salvar_dados()

    atualizar_dashboard()
    atualizar_lista_produtos()
    atualizar_combo_produtos()
    atualizar_estoque()
    atualizar_historico()

    messagebox.showinfo(
        "Venda excluída",
        "A venda foi excluída com sucesso e o estoque foi atualizado."
    )

def editar_venda():
    selecionado = tree_historico.selection()

    if not selecionado:
        messagebox.showwarning(
            "Editar venda",
            "Selecione uma venda no histórico primeiro."
        )
        return

    item = selecionado[0]

    try:
        indice = int(item)
        venda = vendas[indice]
    except (ValueError, IndexError):
        messagebox.showerror(
            "Erro",
            "Não foi possível identificar a venda selecionada."
        )
        return

    produto = venda["produto"]
    quantidade_antiga = venda["quantidade"]

    janela_edicao = tk.Toplevel(janela)
    janela_edicao.title("Editar Venda")
    janela_edicao.geometry("350x250")
    janela_edicao.resizable(False, False)
    janela_edicao.configure(bg="#F4F6F8")

    tk.Label(
        janela_edicao,
        text="EDITAR VENDA",
        bg="#F4F6F8",
        font=("Segoe UI", 14, "bold")
    ).pack(pady=(15, 10))

    tk.Label(
        janela_edicao,
        text=f"Produto: {produto}",
        bg="#F4F6F8",
        font=("Segoe UI", 10)
    ).pack(pady=5)

    tk.Label(
        janela_edicao,
        text=f"Quantidade atual: {quantidade_antiga}",
        bg="#F4F6F8",
        font=("Segoe UI", 10)
    ).pack(pady=5)

    tk.Label(
        janela_edicao,
        text="Nova quantidade:",
        bg="#F4F6F8",
        font=("Segoe UI", 10, "bold")
    ).pack(pady=(10, 3))

    entry_nova_quantidade = tk.Entry(
        janela_edicao,
        width=15,
        justify="center",
        font=("Segoe UI", 11)
    )
    entry_nova_quantidade.pack()

    entry_nova_quantidade.insert(0, str(quantidade_antiga))
    entry_nova_quantidade.focus()
    entry_nova_quantidade.select_range(0, tk.END)

    def salvar_edicao():
        try:
            nova_quantidade = int(entry_nova_quantidade.get())
        except ValueError:
            messagebox.showerror(
                "Quantidade inválida",
                "Digite uma quantidade inteira válida."
            )
            return

        if nova_quantidade <= 0:
            messagebox.showwarning(
                "Quantidade inválida",
                "A quantidade deve ser maior que zero."
            )
            return

        diferenca = nova_quantidade - quantidade_antiga

        if produto not in produtos:
            messagebox.showerror(
                "Erro",
                "O produto não foi encontrado no estoque."
            )
            return

        estoque_atual = produtos[produto]["estoque"]

        if diferenca > 0 and estoque_atual < diferenca:
            messagebox.showwarning(
                "Estoque insuficiente",
                f"Não há estoque suficiente para aumentar a venda.\n\n"
                f"Estoque disponível: {estoque_atual}\n"
                f"Quantidade adicional necessária: {diferenca}"
            )
            return

        # Ajusta o estoque
        produtos[produto]["estoque"] -= diferenca

        # Atualiza a quantidade da venda
        venda["quantidade"] = nova_quantidade

        # Recalcula os valores
        venda["faturamento"] = venda["preco"] * nova_quantidade
        venda["custo_total"] = venda["custo"] * nova_quantidade
        venda["lucro"] = (
            venda["faturamento"] - venda["custo_total"]
        )

        salvar_dados()

        atualizar_dashboard()
        atualizar_lista_produtos()
        atualizar_combo_produtos()
        atualizar_estoque()
        atualizar_historico()

        janela_edicao.destroy()

        messagebox.showinfo(
            "Venda atualizada",
            "A venda foi atualizada com sucesso e o estoque foi ajustado."
        )
    tk.Button(
        janela_edicao,
        text="💾 SALVAR ALTERAÇÃO",
        command=salvar_edicao,
        bg="#27AE60",
        fg="white",
        font=("Segoe UI", 10, "bold"),
        padx=15,
        pady=8,
        cursor="hand2"
    ).pack(pady=15)

def abrir_relatorio_financeiro():
    faturamento = sum(v["faturamento"] for v in vendas)
    custos = sum(v["custo_total"] for v in vendas)
    lucro = sum(v["lucro"] for v in vendas)
    quantidade = sum(v["quantidade"] for v in vendas)

    # Produto mais vendido
    vendas_por_produto = {}

    for v in vendas:
        produto = v["produto"]
        vendas_por_produto[produto] = (
            vendas_por_produto.get(produto, 0) + v["quantidade"]
        )

    if vendas_por_produto:
        produto_mais_vendido = max(
            vendas_por_produto,
            key=vendas_por_produto.get
        )
        quantidade_produto = vendas_por_produto[produto_mais_vendido]
    else:
        produto_mais_vendido = "Nenhuma venda"
        quantidade_produto = 0

    # Produtos com estoque baixo
    produtos_baixo = []

    for nome, dados in produtos.items():
        if dados["estoque"] <= dados["estoque_minimo"]:
            produtos_baixo.append(
                f"{nome} — {dados['estoque']} unidade(s)"
            )

    janela_relatorio = tk.Toplevel(janela)
    janela_relatorio.title("Relatório Financeiro")
    janela_relatorio.geometry("700x600")
    janela_relatorio.resizable(False, False)
    janela_relatorio.configure(bg="#F4F6F8")

    tk.Label(
        janela_relatorio,
        text="📊 RELATÓRIO FINANCEIRO",
        bg="#17202A",
        fg="white",
        font=("Segoe UI", 18, "bold"),
        pady=15
    ).pack(fill="x")

    # Resumo financeiro
    frame_resumo = tk.Frame(
        janela_relatorio,
        bg="#F4F6F8"
    )

    frame_resumo.pack(
        fill="x",
        padx=20,
        pady=20
    )

    def criar_indicador(parent, titulo, valor, coluna):
        frame = tk.Frame(
            parent,
            bg="white",
            bd=1,
            relief="solid"
        )

        frame.grid(
            row=0,
            column=coluna,
            padx=5,
            sticky="nsew"
        )

        parent.grid_columnconfigure(
            coluna,
            weight=1
        )

        tk.Label(
            frame,
            text=titulo,
            bg="white",
            fg="#566573",
            font=("Segoe UI", 9, "bold")
        ).pack(
            pady=(12, 4)
        )

        tk.Label(
            frame,
            text=valor,
            bg="white",
            fg="#17202A",
            font=("Segoe UI", 14, "bold")
        ).pack(
            pady=(0, 12)
        )

    criar_indicador(
        frame_resumo,
        "FATURAMENTO",
        moeda(faturamento),
        0
    )

    criar_indicador(
        frame_resumo,
        "CUSTOS",
        moeda(custos),
        1
    )

    criar_indicador(
        frame_resumo,
        "LUCRO",
        moeda(lucro),
        2
    )

    # Informações gerais
    frame_informacoes = tk.Frame(
        janela_relatorio,
        bg="white",
        bd=1,
        relief="solid"
    )

    frame_informacoes.pack(
        fill="x",
        padx=25,
        pady=5
    )

    tk.Label(
        frame_informacoes,
        text="RESUMO DAS VENDAS",
        bg="white",
        fg="#17202A",
        font=("Segoe UI", 13, "bold")
    ).pack(
        anchor="w",
        padx=20,
        pady=(15, 10)
    )

    informacoes = [
        f"📦 Itens vendidos: {quantidade}",
        f"🧾 Número de vendas: {len(vendas)}",
        f"🏆 Produto mais vendido: {produto_mais_vendido}",
        f"🔢 Quantidade vendida: {quantidade_produto}"
    ]

    for texto in informacoes:
        tk.Label(
            frame_informacoes,
            text=texto,
            bg="white",
            fg="#34495E",
            font=("Segoe UI", 10)
        ).pack(
            anchor="w",
            padx=20,
            pady=4
        )

    # Estoque baixo
    frame_estoque_baixo = tk.Frame(
        janela_relatorio,
        bg="white",
        bd=1,
        relief="solid"
    )

    frame_estoque_baixo.pack(
        fill="both",
        expand=True,
        padx=25,
        pady=15
    )

    tk.Label(
        frame_estoque_baixo,
        text="⚠️ ESTOQUE BAIXO",
        bg="white",
        fg="#C0392B",
        font=("Segoe UI", 13, "bold")
    ).pack(
        anchor="w",
        padx=20,
        pady=(15, 8)
    )

    if produtos_baixo:
        for texto in produtos_baixo:
            tk.Label(
                frame_estoque_baixo,
                text=texto,
                bg="white",
                fg="#34495E",
                font=("Segoe UI", 10)
            ).pack(
                anchor="w",
                padx=20,
                pady=3
            )
    else:
        tk.Label(
            frame_estoque_baixo,
            text="Nenhum produto está com estoque baixo.",
            bg="white",
            fg="#27AE60",
            font=("Segoe UI", 10, "bold")
        ).pack(
            anchor="w",
            padx=20,
            pady=3
        )

    tk.Button(
        janela_relatorio,
        text="FECHAR",
        command=janela_relatorio.destroy,
        bg="#566573",
        fg="white",
        font=("Segoe UI", 10, "bold"),
        padx=25,
        pady=8,
        cursor="hand2"
    ).pack(
        pady=(0, 15)
    )

def limpar_busca():
    entry_busca.delete(0, tk.END)
    atualizar_historico()

def gerar_excel():
    if not produtos and not vendas:
        messagebox.showwarning(
            "Sem dados",
            "Ainda não existem dados para gerar o relatório."
        )
        return

    wb = Workbook()

    # ================= PRODUTOS =================

    ws = wb.active
    ws.title = "Produtos"

    ws.append([
        "Produto",
        "Preço de Venda",
        "Custo",
        "Estoque Inicial",
        "Estoque Atual",
        "Estoque Mínimo",
        "Status"
    ])

    for nome, dados in sorted(produtos.items()):
        atual = dados["estoque"]
        minimo = dados["estoque_minimo"]

        if atual <= 0:
            status = "ESGOTADO"
        elif atual <= minimo:
            status = "ESTOQUE BAIXO"
        else:
            status = "OK"

        ws.append([
            nome,
            dados["preco"],
            dados["custo"],
            dados["estoque_inicial"],
            atual,
            minimo,
            status
        ])

    # ================= VENDAS =================

    ws2 = wb.create_sheet("Vendas")

    ws2.append([
        "Data/Hora",
        "Produto",
        "Preço",
        "Custo",
        "Quantidade",
        "Faturamento",
        "Custo Total",
        "Lucro"
    ])

    for v in vendas:
        ws2.append([
            v.get("data", ""),
            v["produto"],
            v["preco"],
            v["custo"],
            v["quantidade"],
            v["faturamento"],
            v["custo_total"],
            v["lucro"]
        ])

    # ================= DASHBOARD =================

    ws3 = wb.create_sheet("Dashboard")

    faturamento = sum(v["faturamento"] for v in vendas)
    custos = sum(v["custo_total"] for v in vendas)
    lucro = sum(v["lucro"] for v in vendas)
    quantidade = sum(v["quantidade"] for v in vendas)

    ws3["A1"] = "DASHBOARD FINANCEIRO"
    ws3["A1"].font = Font(
        bold=True,
        size=16
    )

    ws3["A3"] = "Indicador"
    ws3["B3"] = "Valor"

    ws3["A4"] = "Faturamento"
    ws3["B4"] = faturamento

    ws3["A5"] = "Custos"
    ws3["B5"] = custos

    ws3["A6"] = "Lucro"
    ws3["B6"] = lucro

    ws3["A7"] = "Itens vendidos"
    ws3["B7"] = quantidade

    ws3["A8"] = "Número de vendas"
    ws3["B8"] = len(vendas)

    ws3["A9"] = "Produtos cadastrados"
    ws3["B9"] = len(produtos)

    # ================= PRODUTO MAIS VENDIDO =================

    vendas_por_produto = {}

    for v in vendas:
        nome = v["produto"]

        vendas_por_produto[nome] = (
            vendas_por_produto.get(nome, 0)
            + v["quantidade"]
        )

    if vendas_por_produto:
        produto_mais_vendido = max(
            vendas_por_produto,
            key=vendas_por_produto.get
        )

        quantidade_mais_vendido = vendas_por_produto[
            produto_mais_vendido
        ]
    else:
        produto_mais_vendido = "Nenhuma venda"
        quantidade_mais_vendido = 0

    ws3["A11"] = "Produto mais vendido"
    ws3["B11"] = produto_mais_vendido

    ws3["A12"] = "Quantidade vendida"
    ws3["B12"] = quantidade_mais_vendido

    # ================= FORMATAÇÃO =================

    for sheet in wb.worksheets:

        # Cabeçalho
        for cell in sheet[1]:
            cell.font = Font(bold=True)
            cell.fill = PatternFill(
                "solid",
                fgColor="D9EAF7"
            )
            cell.alignment = Alignment(
                horizontal="center"
            )

        # Largura das colunas
        for col in sheet.columns:
            letra = col[0].column_letter

            maior = max(
                len(str(c.value or ""))
                for c in col
            )

            sheet.column_dimensions[
                letra
            ].width = min(maior + 3, 30)

    # Formatação especial do Dashboard
    for cell in ws3[3]:
        cell.font = Font(bold=True)
        cell.fill = PatternFill(
            "solid",
            fgColor="D9EAF7"
        )
        cell.alignment = Alignment(
            horizontal="center"
        )

    # Formatação de valores financeiros
    for sheet in [ws, ws2, ws3]:

        for row in sheet.iter_rows():

            for cell in row:

                if isinstance(cell.value, float):

                    cell.number_format = (
                        'R$ #,##0.00'
                    )

    # Congelar cabeçalhos
    ws.freeze_panes = "A2"
    ws2.freeze_panes = "A2"

    # Filtro nas tabelas
    ws.auto_filter.ref = ws.dimensions
    ws2.auto_filter.ref = ws2.dimensions

    try:

        wb.save(r"C:\Users\Jhow\Downloads\relatorio_vendas.xlsx")

    except PermissionError:

        messagebox.showerror(
            "Erro ao salvar Excel",
            "O arquivo relatorio_vendas.xlsx está aberto.\n\n"
            "Feche o arquivo no Excel e tente novamente."
        )

        return

    messagebox.showinfo(
        "Excel",
        "Relatório criado com sucesso!\n\n"
        "O arquivo contém:\n"
        "• Produtos\n"
        "• Vendas\n"
        "• Dashboard financeiro"
    )


# ================= INTERFACE =================

janela = tk.Tk()
janela.title("Sistema de Vendas")
janela.geometry("1150x760")
janela.minsize(1000, 680)
janela.configure(bg="#F4F6F8")

fonte_normal = ("Segoe UI", 10)
fonte_titulo = ("Segoe UI", 22, "bold")

style = ttk.Style()
try:
    style.theme_use("clam")
except tk.TclError:
    pass

style.configure(
    "TNotebook",
    background="#F4F6F8",
    borderwidth=0
)

style.configure(
    "TNotebook.Tab",
    padding=(18, 10),
    font=("Segoe UI", 10, "bold")
)

style.configure(
    "Treeview",
    rowheight=30,
    font=("Segoe UI", 9)
)

style.configure(
    "Treeview.Heading",
    font=("Segoe UI", 9, "bold")
)

# Cabeçalho
topo = tk.Frame(janela, bg="#17202A", height=82)
topo.pack(fill="x")
topo.pack_propagate(False)

tk.Label(
    topo,
    text="SISTEMA DE VENDAS",
    bg="#17202A",
    fg="white",
    font=fonte_titulo
).pack(side="left", padx=25, pady=13)

tk.Label(
    topo,
    text="Produtos • Vendas • Estoque • Lucro",
    bg="#17202A",
    fg="#BFC9CA",
    font=("Segoe UI", 10)
).place(x=28, y=52)

# Dashboard
dashboard = tk.Frame(janela, bg="#F4F6F8")
dashboard.pack(fill="x", padx=20, pady=18)

def criar_card(parent, titulo, valor, coluna):
    frame = tk.Frame(
        parent,
        bg="white",
        bd=1,
        relief="solid"
    )

    frame.grid(
        row=0,
        column=coluna,
        padx=5,
        sticky="nsew"
    )

    parent.grid_columnconfigure(coluna, weight=1)

    tk.Label(
        frame,
        text=titulo,
        bg="white",
        fg="#566573",
        font=("Segoe UI", 9, "bold")
    ).pack(
        anchor="w",
        padx=12,
        pady=(11, 2)
    )

    label = tk.Label(
        frame,
        text=valor,
        bg="white",
        fg="#17202A",
        font=("Segoe UI", 15, "bold")
    )

    label.pack(
        anchor="w",
        padx=12,
        pady=(0, 11)
    )

    return label


lbl_faturamento = criar_card(
    dashboard, "FATURAMENTO", "R$ 0,00", 0
)

lbl_custos = criar_card(
    dashboard, "CUSTOS", "R$ 0,00", 1
)

lbl_lucro = criar_card(
    dashboard, "LUCRO", "R$ 0,00", 2
)

lbl_itens = criar_card(
    dashboard, "ITENS VENDIDOS", "0", 3
)

lbl_vendas = criar_card(
    dashboard, "VENDAS", "0", 4
)

lbl_produtos = criar_card(
    dashboard, "PRODUTOS", "0", 5
)

notebook = ttk.Notebook(janela)
notebook.pack(
    fill="both",
    expand=True,
    padx=20,
    pady=(0, 15)
)

# ================= ABA PRODUTOS =================

aba_produtos = tk.Frame(
    notebook,
    bg="#F4F6F8"
)

notebook.add(
    aba_produtos,
    text="  Produtos  "
)

form_produto = tk.Frame(
    aba_produtos,
    bg="white",
    bd=1,
    relief="solid"
)

form_produto.pack(
    fill="x",
    padx=10,
    pady=10
)

tk.Label(
    form_produto,
    text="Cadastrar produto",
    bg="white",
    fg="#17202A",
    font=("Segoe UI", 15, "bold")
).grid(
    row=0,
    column=0,
    columnspan=4,
    sticky="w",
    padx=20,
    pady=(17, 13)
)

entry_nome_produto = tk.Entry(
    form_produto,
    font=fonte_normal
)

entry_preco_produto = tk.Entry(
    form_produto,
    font=fonte_normal
)

entry_custo_produto = tk.Entry(
    form_produto,
    font=fonte_normal
)

entry_estoque_produto = tk.Entry(
    form_produto,
    font=fonte_normal
)

entry_minimo_produto = tk.Entry(
    form_produto,
    font=fonte_normal
)

campos_produto = [
    ("Nome do produto", entry_nome_produto, 1, 0),
    ("Preço de venda", entry_preco_produto, 1, 2),
    ("Custo", entry_custo_produto, 2, 0),
    ("Estoque", entry_estoque_produto, 2, 2),
    ("Estoque mínimo", entry_minimo_produto, 3, 0),
]

for texto, entrada, linha, coluna in campos_produto:
    tk.Label(
        form_produto,
        text=texto,
        bg="white",
        fg="#34495E",
        font=("Segoe UI", 9, "bold")
    ).grid(
        row=linha,
        column=coluna,
        sticky="w",
        padx=(20, 8),
        pady=7
    )

    entrada.grid(
        row=linha,
        column=coluna + 1,
        sticky="ew",
        padx=(0, 20),
        pady=7
    )

for col in range(4):
    form_produto.grid_columnconfigure(
        col,
        weight=1
    )

botoes_produto = tk.Frame(
    form_produto,
    bg="white"
)

botoes_produto.grid(
    row=4,
    column=0,
    columnspan=4,
    sticky="w",
    padx=20,
    pady=18
)

tk.Button(
    botoes_produto,
    text="SALVAR PRODUTO",
    command=cadastrar_produto,
    bg="#1ABC9C",
    fg="white",
    activebackground="#16A085",
    activeforeground="white",
    font=("Segoe UI", 10, "bold"),
    relief="flat",
    padx=18,
    pady=9,
    cursor="hand2"
).pack(
    side="left",
    padx=(0, 10)
)

tk.Button(
    botoes_produto,
    text="LIMPAR",
    command=limpar_formulario_produto,
    bg="#566573",
    fg="white",
    activebackground="#34495E",
    activeforeground="white",
    font=("Segoe UI", 10, "bold"),
    relief="flat",
    padx=18,
    pady=9,
    cursor="hand2"
).pack(
    side="left"
)

tk.Label(
    aba_produtos,
    text="Dica: clique em um produto da tabela para carregar seus dados no formulário.",
    bg="#F4F6F8",
    fg="#7B7D7D",
    font=("Segoe UI", 9)
).pack(
    anchor="w",
    padx=10,
    pady=(0, 6)
)

frame_produtos = tk.Frame(aba_produtos)
frame_produtos.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=5
)

tree_produtos = ttk.Treeview(
    frame_produtos,
    columns=(
        "produto",
        "preco",
        "custo",
        "estoque",
        "minimo",
        "status"
    ),
    show="headings"
)

for col, texto, largura in [
    ("produto", "Produto", 250),
    ("preco", "Preço", 120),
    ("custo", "Custo", 120),
    ("estoque", "Estoque", 100),
    ("minimo", "Mínimo", 100),
    ("status", "Status", 150)
]:
    tree_produtos.heading(
        col,
        text=texto
    )

    tree_produtos.column(
        col,
        width=largura,
        anchor="center" if col != "produto" else "w"
    )

tree_produtos.pack(
    fill="both",
    expand=True
)

tree_produtos.bind(
    "<Double-1>",
    selecionar_produto
)

# ================= ABA VENDAS =================

aba_venda = tk.Frame(
    notebook,
    bg="#F4F6F8"
)

notebook.add(
    aba_venda,
    text="  Registrar Venda  "
)

form_venda = tk.Frame(
    aba_venda,
    bg="white",
    bd=1,
    relief="solid"
)

form_venda.pack(
    fill="x",
    padx=10,
    pady=10
)

tk.Label(
    form_venda,
    text="Nova venda",
    bg="white",
    fg="#17202A",
    font=("Segoe UI", 15, "bold")
).grid(
    row=0,
    column=0,
    columnspan=4,
    sticky="w",
    padx=20,
    pady=(18, 13)
)

tk.Label(
    form_venda,
    text="Produto",
    bg="white",
    fg="#34495E",
    font=("Segoe UI", 9, "bold")
).grid(
    row=1,
    column=0,
    sticky="w",
    padx=(20, 8),
    pady=7
)

combo_produto = ttk.Combobox(
    form_venda,
    state="readonly",
    font=fonte_normal
)

combo_produto.grid(
    row=1,
    column=1,
    sticky="ew",
    padx=(0, 20),
    pady=7
)

combo_produto.bind(
    "<<ComboboxSelected>>",
    atualizar_dados_venda
)

tk.Label(
    form_venda,
    text="Quantidade",
    bg="white",
    fg="#34495E",
    font=("Segoe UI", 9, "bold")
).grid(
    row=1,
    column=2,
    sticky="w",
    padx=(20, 8),
    pady=7
)

entry_quantidade_venda = tk.Entry(
    form_venda,
    font=fonte_normal
)

entry_quantidade_venda.insert(
    0,
    "1"
)

entry_quantidade_venda.grid(
    row=1,
    column=3,
    sticky="ew",
    padx=(0, 20),
    pady=7
)

tk.Label(
    form_venda,
    text="Preço de venda",
    bg="white",
    fg="#34495E",
    font=("Segoe UI", 9, "bold")
).grid(
    row=2,
    column=0,
    sticky="w",
    padx=(20, 8),
    pady=7
)

entry_preco_venda = tk.Entry(
    form_venda,
    font=fonte_normal,
    state="readonly"
)

entry_preco_venda.grid(
    row=2,
    column=1,
    sticky="ew",
    padx=(0, 20),
    pady=7
)

tk.Label(
    form_venda,
    text="Custo",
    bg="white",
    fg="#34495E",
    font=("Segoe UI", 9, "bold")
).grid(
    row=2,
    column=2,
    sticky="w",
    padx=(20, 8),
    pady=7
)

entry_custo_venda = tk.Entry(
    form_venda,
    font=fonte_normal,
    state="readonly"
)

entry_custo_venda.grid(
    row=2,
    column=3,
    sticky="ew",
    padx=(0, 20),
    pady=7
)

lbl_estoque_disponivel = tk.Label(
    form_venda,
    text="Estoque disponível: -",
    bg="white",
    fg="#2471A3",
    font=("Segoe UI", 10, "bold")
)

lbl_estoque_disponivel.grid(
    row=3,
    column=0,
    columnspan=2,
    sticky="w",
    padx=20,
    pady=8
)

for col in range(4):
    form_venda.grid_columnconfigure(
        col,
        weight=1
    )

tk.Button(
    form_venda,
    text="REGISTRAR VENDA",
    command=registrar_venda,
    bg="#1ABC9C",
    fg="white",
    activebackground="#16A085",
    activeforeground="white",
    font=("Segoe UI", 10, "bold"),
    relief="flat",
    padx=18,
    pady=10,
    cursor="hand2"
).grid(
    row=4,
    column=0,
    columnspan=4,
    sticky="w",
    padx=20,
    pady=(8, 20)
)

# ================= ABA ESTOQUE =================

aba_estoque = tk.Frame(
    notebook,
    bg="#F4F6F8"
)

notebook.add(
    aba_estoque,
    text="  Estoque  "
)

tk.Label(
    aba_estoque,
    text="Controle de estoque",
    bg="#F4F6F8",
    fg="#17202A",
    font=("Segoe UI", 16, "bold")
).pack(
    anchor="w",
    padx=10,
    pady=(12, 8)
)

frame_estoque = tk.Frame(
    aba_estoque
)

frame_estoque.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=5
)

tree_estoque = ttk.Treeview(
    frame_estoque,
    columns=(
        "produto",
        "inicial",
        "vendido",
        "atual",
        "minimo",
        "status"
    ),
    show="headings"
)

for col, texto, largura in [
    ("produto", "Produto", 240),
    ("inicial", "Estoque Inicial", 130),
    ("vendido", "Vendido", 110),
    ("atual", "Atual", 110),
    ("minimo", "Mínimo", 100),
    ("status", "Status", 150)
]:
    tree_estoque.heading(
        col,
        text=texto
    )

    tree_estoque.column(
        col,
        width=largura,
        anchor="center" if col != "produto" else "w"
    )

tree_estoque.pack(
    fill="both",
    expand=True
)

# ================= ABA HISTÓRICO =================

aba_historico = tk.Frame(
    notebook,
    bg="#F4F6F8"
)

notebook.add(
    aba_historico,
    text="  Histórico de Vendas  "
)

barra_busca = tk.Frame(
    aba_historico,
    bg="#F4F6F8"
)

barra_busca.pack(
    fill="x",
    padx=10,
    pady=(12, 8)
)

tk.Label(
    barra_busca,
    text="Pesquisar produto:",
    bg="#F4F6F8",
    fg="#34495E",
    font=("Segoe UI", 10, "bold")
).pack(
    side="left"
)

entry_busca = tk.Entry(
    barra_busca,
    font=fonte_normal,
    width=32
)

entry_busca.pack(
    side="left",
    padx=8
)

entry_busca.bind(
    "<KeyRelease>",
    lambda event: atualizar_historico()
)

tk.Button(
    barra_busca,
    text="LIMPAR",
    command=limpar_busca,
    bg="#566573",
    fg="white",
    relief="flat",
    font=("Segoe UI", 9, "bold"),
    padx=12,
    pady=5
).pack(
    side="left"
)

tk.Button(
    barra_busca,
    text="GERAR EXCEL",
    command=gerar_excel,
    bg="#2471A3",
    fg="white",
    relief="flat",
    font=("Segoe UI", 9, "bold"),
    padx=12,
    pady=5
).pack(
    side="right"
)

frame_historico = tk.Frame(
    aba_historico
)

frame_historico.pack(
    fill="both",
    expand=True,
    padx=10,
    pady=5
)

tree_historico = ttk.Treeview(
    frame_historico,
    columns=(
        "data",
        "produto",
        "preco",
        "custo",
        "quantidade",
        "faturamento",
        "lucro"
    ),
    show="headings"
)

for col, texto, largura in [
    ("data", "Data/Hora", 160),
    ("produto", "Produto", 210),
    ("preco", "Preço", 110),
    ("custo", "Custo", 110),
    ("quantidade", "Quantidade", 100),
    ("faturamento", "Faturamento", 130),
    ("lucro", "Lucro", 130)
]:
    tree_historico.heading(
        col,
        text=texto
    )

    tree_historico.column(
        col,
        width=largura,
        anchor="center" if col != "produto" else "w"
    )

tree_historico.pack(
    fill="both",
    expand=True
)

tree_historico.pack(
    fill="both",
    expand=True
)
# Botões de ações do histórico
frame_botoes_historico = tk.Frame(
    frame_historico,
    bg="#F4F6F8"
)

frame_botoes_historico.pack(pady=10)

# Botão Editar
tk.Button(
    frame_botoes_historico,
    text="✏️ EDITAR VENDA",
    command=editar_venda,
    bg="#2980B9",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    padx=20,
    pady=8,
    cursor="hand2"
).pack(
    side="left",
    padx=5
)

# Botão Excluir
tk.Button(
    frame_botoes_historico,
    text="🗑️ EXCLUIR VENDA",
    command=excluir_venda,
    bg="#C0392B",
    fg="white",
    font=("Segoe UI", 10, "bold"),
    padx=20,
    pady=8,
    cursor="hand2"
).pack(
    side="left",
    padx=5
)

tk.Button(
    barra_busca,
    text="GERAR EXCEL",
    command=gerar_excel,
    bg="#2471A3",
    fg="white",
    relief="flat",
    font=("Segoe UI", 9, "bold"),
    padx=12,
    pady=5
).pack(
    side="right"
)

tk.Button(
    barra_busca,
    text="📊 RELATÓRIO DE VENDAS",
    command=abrir_relatorio_financeiro,
    bg="#8E44AD",
    fg="white",
    relief="flat",
    font=("Segoe UI", 9, "bold"),
    padx=12,
    pady=5,
    cursor="hand2"
).pack(
    side="right",
    padx=(0, 8)
)

# Rodapé
tk.Label(
    janela,
    text="Dados salvos automaticamente em dados_sistema.xlsx",
    bg="#F4F6F8",
    fg="#7B7D7D",
    font=("Segoe UI", 8)
).pack(
    pady=(0, 7)
)

carregar_dados()

entry_nome_produto.focus()

janela.mainloop()
