# DFC + Dashboard a partir do Razão Contábil (MV Soul)

> Projeto de demonstração. O documento de exemplo foi **inspirado em um documento de uma organização real, mas todos os dados são fictícios**.

## Teste online (sem instalar nada)
https://claude.ai/artifact/4z8PyXeFSiaygc82htFAW3
Abra o link, envie o Razão Contábil (PDF ou TXT) e o painel é montado no seu aparelho. Os arquivos não saem dele.

## Arquivos
| Arquivo | Função |
|---|---|
| `template.html` | Painel (tabela DFC, cartões, gráficos). Funciona sozinho no navegador ou recebe os dados do Python |
| `dfc_core.py` | Leitura do PDF e soma por conta/mês. Também gera `DFC.xlsx` e `Dashboard.png` (`python dfc_core.py Razao.pdf`) |
| `gerar_dashboard.py` | Lê o PDF com Python e grava o `dashboard.html` pronto |
| `exemplo/` | Razão fictício e o dashboard gerado com ele |

## Uso no PyCharm
1. `pip install -r requirements.txt`
2. `python gerar_dashboard.py exemplo/Razao_Contabil_Base_Ficticia_SUS.pdf --empresa "Minha Empresa" --saldo-inicial 50000 --base passivo`
3. Abra o `dashboard.html` no navegador.

## Atenção
Sem contas de Caixa/Bancos (1.1.1) no Razão, o DFC é uma **estimativa**: recebimentos = créditos das receitas (competência). Informe o saldo inicial real.
