"""
DFC (método direto) a partir do Razão Contábil do MV Soul
=========================================================
Lê o PDF (ou TXT) do Razão Contábil, classifica as contas e gera:
  - saida_dfc/DFC.xlsx       -> DFC mensal (planilha formatada)
  - saida_dfc/Dashboard.png  -> painel com indicadores e gráficos

Instalação (Terminal do PyCharm):
    pip install pymupdf pandas matplotlib openpyxl

Uso:
    1) Ajuste as constantes da seção CONFIGURAÇÃO e clique em Run; ou
    2) python dfc_mv_soul.py Razao.pdf --saldo-inicial 50000 --base passivo

ATENÇÃO: se o Razão não tiver contas de Caixa/Bancos (1.1.1), o DFC é uma
ESTIMATIVA: recebimentos = créditos das receitas (competência).
"""
import argparse
import re
from collections import defaultdict
from pathlib import Path

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd

# ------------------------------ CONFIGURAÇÃO ------------------------------
ARQUIVO_ENTRADA = r"Razao_Contabil.pdf"   # PDF ou TXT (pdftotext -layout)
SALDO_INICIAL = 0.0                       # saldo real de caixa em 01/01
BASE_PAGAMENTOS = "passivo"               # "passivo" ou "despesas"
PASTA_SAIDA = "saida_dfc"
EMPRESA = "Empresa"                       # título exibido no painel

# Classificação: (rótulo, regex do código da conta)
RECEBIMENTOS = [  # usa CRÉDITOS
    ("Receitas SUS e incentivos", r"^4\.1\.1"),
    ("Serviços e vendas", r"^4\.1\.2"),
    ("Reembolsos e recuperações", r"^4\.1\.3"),
    ("Outras receitas operacionais", r"^4\.1\.[46]"),
    ("Receitas financeiras", r"^4\.2"),
]
PAGAMENTOS_DESPESAS = [  # usa DÉBITOS das despesas (competência)
    ("Pessoal e encargos", r"^3\.1\.1"),
    ("Serviços de terceiros", r"^3\.1\.2"),
    ("Materiais e suprimentos", r"^3\.1\.3"),
    ("Despesas gerais e administrativas", r"^3\.1\.4\.(01|04)"),
    ("Tributos", r"^3\.1\.5"),
    ("Despesas financeiras", r"^3\.2"),
]
PAGAMENTOS_PASSIVO = [  # usa DÉBITOS dos passivos operacionais (pagamentos)
    ("Fornecedores", r"^2\.1\.1\.0[23]"),
    ("Pessoal e encargos a pagar", r"^2\.1\.1\.05"),
    ("Tributos e contribuições", r"^2\.1\.1\.(06|09)"),
    ("Parcelamentos e passivos trabalhistas",
     r"^(2\.1\.1\.(04|07|08|15|17)|2\.2\.1\.)"),
]
INVESTIMENTO = [("Convênios e investimentos (líquido)", r"^2\.1\.1\.1[34]")]
EMPRESTIMOS = r"^2\.1\.1\.1[12]"
# --------------------------------------------------------------------------

RE_CONTA = re.compile(
    r"Conta Anal[ií]tica:\s+\d+\s+([\d.]+)\s+-\s+(.*?)\s+Saldo Anterior:")
RE_LANC = re.compile(
    r"^\s*(\d{2}/\d{2}/\d{4})?\s*(\d+)\s+(?:(\d+)\s+)?"
    r"(-?[\d.]+,\d\d)\s+(-?[\d.]+,\d\d)\s+(-?[\d.]+,\d\d)\s+(.*)$")
MESES = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun",
         "Jul", "Ago", "Set", "Out", "Nov", "Dez"]


def br_num(txt: str) -> float:
    return float(txt.replace(".", "").replace(",", "."))


def brl(v: float) -> str:
    return ("-" if v < 0 else "") + "R$ " + f"{abs(v):,.0f}".replace(",", ".")


# ------------------------------- LEITURA ---------------------------------
def linhas_do_pdf(caminho: Path):
    """Reconstrói as linhas de texto de cada página agrupando palavras por Y."""
    try:
        import fitz  # PyMuPDF
    except ImportError:  # alternativa: pip install pdfplumber
        import pdfplumber
        with pdfplumber.open(caminho) as pdf:
            for pagina in pdf.pages:
                palavras = sorted(pagina.extract_words(), key=lambda w: (w["bottom"], w["x0"]))
                atual, y = [], None
                for w in palavras:
                    if y is None or abs(w["bottom"] - y) > 2:
                        if atual:
                            yield " ".join(t for _, t in sorted(atual))
                        atual, y = [], w["bottom"]
                    atual.append((w["x0"], w["text"]))
                if atual:
                    yield " ".join(t for _, t in sorted(atual))
        return
    with fitz.open(caminho) as doc:
        total = len(doc)
        for i, pagina in enumerate(doc, 1):
            palavras = sorted(pagina.get_text("words"), key=lambda w: (w[3], w[0]))
            atual, y = [], None
            for w in palavras:
                if y is None or abs(w[3] - y) > 2:
                    if atual:
                        yield " ".join(p[4] for p in sorted(atual))
                    atual, y = [], w[3]
                atual.append((w[0], w[4]))
            if atual:
                yield " ".join(p[4] for p in sorted(atual))
            if i % 100 == 0:
                print(f"  página {i}/{total}")


def ler_lancamentos(caminho: Path) -> pd.DataFrame:
    if caminho.suffix.lower() == ".pdf":
        linhas = linhas_do_pdf(caminho)
    else:
        linhas = caminho.read_text(encoding="utf-8", errors="ignore").splitlines()

    conta = nome = data = None
    movs = defaultdict(lambda: [0.0, 0.0, 0])  # (conta, mes) -> [débito, crédito, qtd]
    nomes, n = {}, 0
    for linha in linhas:
        m = RE_CONTA.search(linha)
        if m:
            if m.group(1) != conta:
                conta, data = m.group(1), None
            nomes[conta] = m.group(2).strip()
            continue
        m = RE_LANC.match(linha)
        if not m or not conta:
            continue
        data = m.group(1) or data
        if not data:
            continue
        mes = f"{data[6:]}-{data[3:5]}"
        movs[(conta, mes)][0] += br_num(m.group(4))
        movs[(conta, mes)][1] += br_num(m.group(5))
        movs[(conta, mes)][2] += 1
        n += 1
    if not n:
        raise SystemExit("Nenhum lançamento reconhecido. Use o Razão Contábil do MV Soul.")
    print(f"✓ {n:,} lançamentos em {len(nomes)} contas.".replace(",", "."))
    df = pd.DataFrame(
        [(c, nomes[c], mes, d, cr, q) for (c, mes), (d, cr, q) in movs.items()],
        columns=["conta", "nome", "mes", "debito", "credito", "qtd"])
    return df


# --------------------------------- DFC -----------------------------------
def serie(df, regex, tipo, sinal, meses):
    """Soma por mês as contas que casam com o regex. tipo: C, D ou N (líquido)."""
    sub = df[df["conta"].str.match(regex)]
    g = sub.groupby("mes")[["debito", "credito"]].sum().reindex(meses, fill_value=0)
    base = {"C": g["credito"], "D": g["debito"], "N": g["credito"] - g["debito"]}[tipo]
    return sinal * base


def montar_dfc(df, saldo_inicial, base):
    meses = sorted(df["mes"].unique())
    pagamentos = PAGAMENTOS_PASSIVO if base == "passivo" else PAGAMENTOS_DESPESAS
    secoes = {
        "Atividades operacionais": (
            [(r, x, "C", 1) for r, x in RECEBIMENTOS]
            + [(r, x, "D", -1) for r, x in pagamentos]),
        "Atividades de investimento": [(r, x, "N", 1) for r, x in INVESTIMENTO],
        "Atividades de financiamento": [
            ("Empréstimos: amortizações", EMPRESTIMOS, "D", -1),
            ("Empréstimos: créditos (captações e encargos)", EMPRESTIMOS, "C", 1)],
    }
    linhas, totais, usadas = [], {}, set()
    for secao, regras in secoes.items():
        linhas.append((secao, None, "secao"))
        soma = pd.Series(0.0, index=meses)
        for rotulo, regex, tipo, sinal in regras:
            s = serie(df, regex, tipo, sinal, meses)
            usadas.update(df.loc[df["conta"].str.match(regex), "conta"])
            if s.abs().sum() == 0:
                continue
            soma += s
            linhas.append((rotulo, s, "item"))
        totais[secao] = soma
        linhas.append((f"Caixa líquido – {secao.lower()}", soma, "total"))
    fco, fci, fcf = totais.values()
    variacao = fco + fci + fcf
    saldo = saldo_inicial + variacao.cumsum()
    linhas += [("Variação líquida de caixa", variacao, "total"),
               ("Saldo final de caixa", saldo, "total")]
    todas = [s for _, s, t in linhas if t == "item"]
    entradas = sum(s[s > 0].sum() for s in todas)
    saidas = -sum(s[s < 0].sum() for s in todas)
    cobertura = (saldo.iloc[-1] / (saidas / len(meses))) if saidas and saldo.iloc[-1] > 0 else None
    kpis = {
        "SALDO INICIAL": saldo_inicial, "SALDO FINAL": saldo.iloc[-1],
        "TOTAL DE ENTRADAS": entradas, "TOTAL DE SAÍDAS": saidas,
        "CAIXA OPERACIONAL (FCO)": fco.sum(), "CAIXA DE INVESTIMENTO (FCI)": fci.sum(),
        "CAIXA DE FINANCIAMENTO (FCF)": fcf.sum(), "COBERTURA DE CAIXA": cobertura,
    }
    return meses, linhas, totais, saldo, kpis, len(set(df["conta"]) - usadas)


def rotulo_mes(ym):
    return f"{MESES[int(ym[5:]) - 1]}/{ym[2:4]}"


# -------------------------------- SAÍDAS ---------------------------------
def salvar_excel(meses, linhas, caminho):
    cab = ["Descrição"] + [rotulo_mes(m) for m in meses] + ["Total"]
    dados = []
    for rot, s, tipo in linhas:
        if s is None:
            dados.append([rot] + [None] * (len(meses) + 1))
        else:
            tot = s.iloc[-1] if rot == "Saldo final de caixa" else s.sum()
            dados.append([rot] + list(s.values) + [tot])
    df = pd.DataFrame(dados, columns=cab)
    with pd.ExcelWriter(caminho, engine="openpyxl") as xl:
        df.to_excel(xl, sheet_name="DFC", index=False)
        ws = xl.sheets["DFC"]
        from openpyxl.styles import Font, PatternFill
        ws.column_dimensions["A"].width = 46
        for c in "BCDEFGHIJKLMNOP":
            ws.column_dimensions[c].width = 13
        for cel in ws[1]:
            cel.font = Font(bold=True, color="FFFFFF")
            cel.fill = PatternFill("solid", fgColor="0B7A53")
        for i, (_, _, tipo) in enumerate(linhas, 2):
            for cel in ws[i]:
                if cel.column > 1:
                    cel.number_format = '#,##0;(#,##0);"-"'
                if tipo in ("secao", "total"):
                    cel.font = Font(bold=True)
                if tipo == "secao":
                    cel.fill = PatternFill("solid", fgColor="E3EFEA")
        ws.freeze_panes = "B2"


def salvar_dashboard(meses, totais, saldo, kpis, caminho, titulo, base, nao_class):
    esmeralda, cinza_escuro, cinza_claro = "#0F9D6B", "#5B6670", "#B9C0C7"
    fig = plt.figure(figsize=(14, 9), facecolor="white")
    fig.text(0.03, 0.955, f"Dashboard de Fluxo de Caixa – {titulo}", fontsize=18,
             fontweight="bold", color=esmeralda)
    fig.text(0.03, 0.925, f"Método direto (estimativa) · pagamentos pela base: {base} · valores em R$",
             fontsize=10, color="#555")
    # cartões
    for i, (nome, v) in enumerate(kpis.items()):
        ax = fig.add_axes([0.03 + (i % 4) * 0.24, 0.78 - (i // 4) * 0.11, 0.22, 0.095])
        ax.set_facecolor("#F1F3F2")
        ax.set_xticks([]); ax.set_yticks([])
        for s in ax.spines.values():
            s.set_color("#D3D8D6")
        if nome == "COBERTURA DE CAIXA":
            txt, cor = ("n/d" if v is None else f"{v:.1f} meses".replace(".", ",")), esmeralda
        else:
            txt, cor = brl(v), ("#4B5563" if v < 0 else "#0B7A53" if "(F" in nome or nome == "SALDO FINAL" else esmeralda)
        ax.text(0.5, 0.72, nome, ha="center", fontsize=8.5, color="#555", transform=ax.transAxes)
        ax.text(0.5, 0.25, txt, ha="center", fontsize=16, fontweight="bold", color=cor,
                transform=ax.transAxes)
    rot = [rotulo_mes(m) for m in meses]
    x = range(len(meses))
    w = 0.27
    ax1 = fig.add_axes([0.06, 0.08, 0.40, 0.42])
    for k, (serie_, cor, lab) in enumerate(zip(totais.values(), (esmeralda, cinza_escuro, cinza_claro),
                                               ("Operacional", "Investimento", "Financiamento"))):
        ax1.bar([i + (k - 1) * w for i in x], serie_.values, w, color=cor, label=lab)
    ax1.axhline(0, color="#999", lw=.8)
    ax1.set_title("Fluxo de caixa por atividade (R$)", fontweight="bold")
    ax1.set_xticks(list(x)); ax1.set_xticklabels(rot, rotation=45, fontsize=8)
    ax1.legend(fontsize=8); ax1.grid(axis="y", alpha=.3)
    ax2 = fig.add_axes([0.55, 0.08, 0.40, 0.42])
    ax2.plot(list(x), saldo.values, marker="o", color="#0F9D6B", lw=2.5)
    ax2.axhline(0, color="#999", lw=.8)
    ax2.set_title("Saldo final de caixa no mês (R$)", fontweight="bold")
    ax2.set_xticks(list(x)); ax2.set_xticklabels(rot, rotation=45, fontsize=8)
    ax2.grid(alpha=.3)
    for ax in (ax1, ax2):
        ax.yaxis.set_major_formatter(lambda v, _: f"{v/1e6:.1f} mi" if abs(v) >= 1e6 else f"{v/1e3:.0f} mil")
    fig.text(0.03, 0.012, f"{nao_class} contas sem classificação (provisões, patrimônio etc.) ficaram fora do DFC. "
             "Sem contas de Caixa/Bancos no Razão, trata-se de estimativa.", fontsize=8, color="#777")
    fig.savefig(caminho, dpi=150)
    plt.close(fig)


def main():
    ap = argparse.ArgumentParser(description="DFC a partir do Razão MV Soul")
    ap.add_argument("arquivo", nargs="?", default=ARQUIVO_ENTRADA)
    ap.add_argument("--saldo-inicial", type=float, default=SALDO_INICIAL)
    ap.add_argument("--base", choices=["passivo", "despesas"], default=BASE_PAGAMENTOS)
    ap.add_argument("--saida", default=PASTA_SAIDA)
    ap.add_argument("--empresa", default=EMPRESA)
    a = ap.parse_args()

    entrada, saida = Path(a.arquivo), Path(a.saida)
    saida.mkdir(parents=True, exist_ok=True)
    print(f"Lendo {entrada} …")
    df = ler_lancamentos(entrada)
    meses, linhas, totais, saldo, kpis, nao_class = montar_dfc(df, a.saldo_inicial, a.base)
    salvar_excel(meses, linhas, saida / "DFC.xlsx")
    salvar_dashboard(meses, totais, saldo, kpis, saida / "Dashboard.png",
                     a.empresa, a.base, nao_class)
    print(f"Pronto! Arquivos em: {saida.resolve()}")
    for k, v in kpis.items():
        print(f"  {k:<30} {'n/d' if v is None else (f'{v:.1f} meses' if 'COBERTURA' in k else brl(v))}")


if __name__ == "__main__":
    main()
