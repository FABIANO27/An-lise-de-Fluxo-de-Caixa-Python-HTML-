"""
Gera o dashboard HTML (DFC + Dashboard) a partir do Razão Contábil do MV Soul.

    python gerar_dashboard.py exemplo/Razao_Contabil_Base_Ficticia_SUS.pdf \
        --empresa "SUS - Base_Fictícia" --saldo-inicial 5000000 --base passivo

O Python lê o PDF e soma os lançamentos por conta/mês (dfc_core.py); o HTML
(template.html) classifica as contas e desenha tabela, cartões e gráficos.
Saída: dashboard.html (arquivo único, abre em qualquer navegador).
"""
import argparse
import json
from pathlib import Path

from dfc_core import ler_lancamentos

AQUI = Path(__file__).parent


def main():
    ap = argparse.ArgumentParser(description="Dashboard DFC a partir do Razão MV Soul")
    ap.add_argument("arquivo", help="Razão Contábil em PDF ou TXT")
    ap.add_argument("--empresa", default="Empresa Exemplo Ltda")
    ap.add_argument("--saldo-inicial", type=float, default=0.0)
    ap.add_argument("--base", choices=["passivo", "despesas"], default="passivo")
    ap.add_argument("--saida", default="dashboard.html")
    a = ap.parse_args()

    df = ler_lancamentos(Path(a.arquivo))
    contas = {}
    for r in df.itertuples():
        c = contas.setdefault(r.conta, {"n": r.nome, "k": 0, "m": {}})
        c["k"] += int(r.qtd)
        c["m"][r.mes] = [round(r.debito, 2), round(r.credito, 2)]
    dados = {"A": contas, "N": int(df["qtd"].sum()), "M": sorted(df["mes"].unique()),
             "empresa": a.empresa, "saldo_inicial": a.saldo_inicial,
             "base": "pas" if a.base == "passivo" else "des"}
    js = json.dumps(dados, ensure_ascii=False).replace("</", "<\\/")
    html = (AQUI / "template.html").read_text(encoding="utf-8")
    html = html.replace("/*__DFC_DATA__*/null", js)
    Path(a.saida).write_text(html, encoding="utf-8")
    print(f"Pronto! Abra {Path(a.saida).resolve()}")


if __name__ == "__main__":
    main()
