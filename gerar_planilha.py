"""Gera a planilha de demonstração do robô de cobrança.

A aba Importacao_Diaria imita o relatório diário de parcelas de um sistema
financeiro (como o Procfy): uma linha por parcela, com cliente, telefone,
cidade, valor, vencimento e situação. As datas são montadas em volta de hoje,
para que sempre existam parcelas vencendo no dia em que a demonstração roda.

Uso:
    python gerar_planilha.py              # cria planilha_cobranca.xlsx
    python gerar_planilha.py --substituir # recria, apagando a anterior
"""

import argparse
import random
import sys
from datetime import date, timedelta
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.worksheet.datavalidation import DataValidation

PLANILHA = Path(__file__).parent / "planilha_cobranca.xlsx"
AZUL = PatternFill("solid", fgColor="1F4E78")
FONTE = "Calibri"

COLUNAS_IMPORTACAO = [
    ("Cliente", 26), ("Telefone", 17), ("Cidade", 16), ("Parcela", 9),
    ("Total de parcelas", 10), ("Valor", 12), ("Vencimento", 13),
    ("Situação", 12), ("Data do pagamento", 14),
]
COLUNAS_HISTORICO = [
    ("Cliente", 26), ("Telefone", 17), ("Cidade", 16), ("Parcela", 9),
    ("Total de parcelas", 10), ("Valor", 12), ("Vencimento", 13),
    ("Situação", 12), ("Data do pagamento", 14), ("Confirmação enviada", 14),
]
COLUNAS_ENVIOS = [
    ("Data e hora", 18), ("Tipo", 14), ("Para", 28), ("Telefone", 17),
    ("Situação", 12), ("Mensagem", 80),
]

# Nomes e números de mentira (DDD válido, mas a demonstração nunca envia nada).
CLIENTES = [
    ("Ana Paula Ribeiro", "41 99812-3301", "Curitiba"),
    ("Carlos Eduardo Lima", "41 99745-1122", "Curitiba"),
    ("Juliana Martins", "41 98877-4410", "Curitiba"),
    ("Roberto Souza", "41 99631-0098", "São José dos Pinhais"),
    ("Fernanda Alves", "41 99102-7745", "São José dos Pinhais"),
    ("Marcos Vinícius Rocha", "41 98450-6631", "Colombo"),
    ("Patrícia Gomes", "41 99984-2207", "Colombo"),
    ("Luiz Fernando Costa", "41 9977-123", "Araucária"),  # telefone incompleto, de propósito
    ("Sandra Regina Pires", "41 99311-8854", "Araucária"),
    ("Thiago Mendes", "41 98236-4419", "Pinhais"),
    ("Camila Ferreira", "41 99570-3362", "Pinhais"),
    ("Eduardo Nogueira", "", "Curitiba"),  # sem telefone, de propósito
]


def cabecalho(aba, colunas):
    for numero, (nome, largura) in enumerate(colunas, start=1):
        celula = aba.cell(1, numero, nome)
        celula.fill = AZUL
        celula.font = Font(name=FONTE, bold=True, color="FFFFFF")
        celula.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        aba.column_dimensions[celula.column_letter].width = largura
    aba.row_dimensions[1].height = 30
    aba.freeze_panes = "A2"


def parcelas_de_mentira(hoje):
    """Monta as parcelas de cada cliente: semanais, com a de hoje em aberto.

    Alguns clientes pagaram a parcela anterior há pouco (o robô manda a
    confirmação com o cronograma) e um deles já pagou a de hoje (o robô não
    cobra quem já pagou). Os pagamentos antigos não recebem confirmação.
    """
    sorteio = random.Random(hoje.toordinal())
    linhas = []
    for indice, (nome, telefone, cidade) in enumerate(CLIENTES):
        total = sorteio.choice([3, 4, 5])
        atual = sorteio.randint(2, total)          # a parcela que vence hoje
        valor = sorteio.choice([89.90, 120.00, 150.00, 199.90, 250.00])
        for parcela in range(1, total + 1):
            vencimento = hoje + timedelta(weeks=parcela - atual)
            if parcela == atual - 1 and indice in (0, 4, 9):
                situacao, pago_em = "Pago", hoje - timedelta(days=1)  # pagou ontem, com atraso
            elif parcela < atual:
                situacao, pago_em = "Pago", vencimento - timedelta(days=sorteio.randint(0, 1))
            elif parcela == atual and indice == 2:
                situacao, pago_em = "Pago", hoje  # pagou adiantado no dia
            else:
                situacao, pago_em = "Em aberto", None
            linhas.append([nome, telefone, cidade, parcela, total, valor,
                           vencimento, situacao, pago_em])
    return linhas


def gerar(caminho, hoje):
    livro = Workbook()

    como_usar = livro.active
    como_usar.title = "Como usar"
    textos = [
        "ROBÔ DE COBRANÇA — DEMONSTRAÇÃO",
        "",
        "1. Todo dia, cole o relatório de parcelas do sistema financeiro na aba Importacao_Diaria.",
        "2. Rode o robô (dois cliques no atalho do seu sistema).",
        "3. O robô cobra no WhatsApp quem tem parcela vencendo HOJE e ainda não pagou.",
        "4. Quem pagou recebe a confirmação com o cronograma das parcelas.",
        "5. A aba Historico_Parcelas guarda todas as parcelas, e a aba Envios, tudo o que foi enviado.",
        "6. No fim, cada grupo de cidade recebe o resumo do dia.",
        "",
        "Nesta demonstração o WhatsApp é SIMULADO: nada é enviado de verdade.",
        "As mensagens aparecem em saida/whatsapp_simulado.html.",
        "Nomes e telefones são de mentira.",
    ]
    for numero, texto in enumerate(textos, start=1):
        como_usar.cell(numero, 1, texto).font = Font(name=FONTE, bold=numero == 1, size=14 if numero == 1 else 11)
    como_usar.column_dimensions["A"].width = 100

    importacao = livro.create_sheet("Importacao_Diaria")
    cabecalho(importacao, COLUNAS_IMPORTACAO)
    for linha in parcelas_de_mentira(hoje):
        importacao.append(linha)
    for linha in importacao.iter_rows(min_row=2):
        linha[5].number_format = '"R$" #,##0.00'
        linha[6].number_format = linha[8].number_format = "DD/MM/YYYY"
        for celula in (linha[3], linha[4], linha[7]):
            celula.alignment = Alignment(horizontal="center")
    situacoes = DataValidation(type="list", formula1='"Pago,Em aberto"', allow_blank=False)
    importacao.add_data_validation(situacoes)
    situacoes.add(f"H2:H{importacao.max_row + 500}")

    cabecalho(livro.create_sheet("Historico_Parcelas"), COLUNAS_HISTORICO)
    cabecalho(livro.create_sheet("Envios"), COLUNAS_ENVIOS)
    livro.active = 1
    livro.save(caminho)


def main():
    argumentos = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    argumentos.add_argument("--substituir", action="store_true",
                            help="recria a planilha mesmo se ela já existir")
    opcoes = argumentos.parse_args()

    if PLANILHA.exists() and not opcoes.substituir:
        print(f"A planilha {PLANILHA.name} já existe. Use --substituir para recriar.")
        return
    try:
        gerar(PLANILHA, date.today())
    except PermissionError:
        sys.exit(f"Não consegui salvar {PLANILHA.name}. Ela está aberta no Excel? Feche e tente de novo.")
    print(f"Planilha criada: {PLANILHA.name}")


if __name__ == "__main__":
    main()
