import pandas as pd
import sqlite3

clientes = pd.read_csv("dados/clientes.csv")
vendas = pd.read_csv("dados/vendas.csv")

conexao = sqlite3.connect("banco.db")

clientes.to_sql("clientes", conexao, if_exists="replace", index=False)
vendas.to_sql("vendas", conexao, if_exists="replace", index=False)

faturamento_produtos = pd.read_sql("""
SELECT
    produto,
    SUM(valor) AS faturamento
FROM vendas
GROUP BY produto
ORDER BY faturamento DESC
""", conexao)

top_clientes = pd.read_sql("""
SELECT
    c.nome,
    SUM(v.valor) AS total_comprado
FROM clientes c
INNER JOIN vendas v
ON c.id_cliente = v.id_cliente
GROUP BY c.nome
ORDER BY total_comprado DESC
""", conexao)

with pd.ExcelWriter("relatorio.xlsx") as writer:
    faturamento_produtos.to_excel(
        writer,
        sheet_name="Produtos",
        index=False
    )

    top_clientes.to_excel(
        writer,
        sheet_name="Clientes",
        index=False
    )

print("Relatório Excel gerado com sucesso!")

conexao.close()