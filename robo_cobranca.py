"""Robô de cobrança: lê o relatório do dia e cobra pelo WhatsApp.

O que ele faz, nesta ordem:
1. Lê a aba Importacao_Diaria (o relatório colado do sistema financeiro).
2. Atualiza a aba Historico_Parcelas com a situação de cada parcela.
3. Manda a confirmação, com o cronograma, para quem pagou nos últimos dias.
4. Cobra quem tem parcela vencendo hoje e ainda não pagou.
5. Manda o resumo do dia para o grupo de cada cidade.
6. Anota tudo na aba Envios da própria planilha.

Uso:
    python robo_cobranca.py --rapido        # demonstração, sem esperar entre mensagens
    python robo_cobranca.py --data 2026-11-10
"""

import argparse
import html
import logging
import os
import random
import re
import sys
import time
import webbrowser
from collections import defaultdict
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from pathlib import Path

from openpyxl import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill

PASTA = Path(__file__).parent
PLANILHA = PASTA / "planilha_cobranca.xlsx"
SAIDA = PASTA / "saida"
CONVERSA_HTML = SAIDA / "whatsapp_simulado.html"

DIAS_PARA_CONFIRMAR = 2      # pagamentos mais antigos que isso não recebem confirmação
LIMITE_POR_HORA = 40         # cuidado contra bloqueio do número
INTERVALO = (20, 45)         # segundos aleatórios entre uma mensagem e outra

VERDE = PatternFill("solid", fgColor="DEF7E5")
AMARELO = PatternFill("solid", fgColor="FFF4CC")
VERMELHO = PatternFill("solid", fgColor="FDE2E2")
FONTE = "Calibri"

log = logging.getLogger("robo")


@dataclass
class Parcela:
    cliente: str
    telefone: str
    cidade: str
    numero: int
    total: int
    valor: float
    vencimento: date
    situacao: str
    pago_em: date | None

    @property
    def paga(self):
        return self.situacao.strip().lower() == "pago"

    @property
    def chave(self):
        return (self.cliente.strip().lower(), self.numero)


# ---------------------------------------------------------------- leitura

def para_data(valor):
    if valor in (None, ""):
        return None
    if isinstance(valor, datetime):
        return valor.date()
    if isinstance(valor, date):
        return valor
    return datetime.strptime(str(valor).strip(), "%d/%m/%Y").date()


def para_valor(valor):
    if isinstance(valor, (int, float)):
        return float(valor)
    texto = str(valor).replace("R$", "").replace(".", "").replace(",", ".").strip()
    return float(texto)


def ler_importacao(aba):
    """Lê o relatório do dia. Linhas com problema vão para o log, sem parar o robô."""
    colunas = {celula.value: indice for indice, celula in enumerate(aba[1]) if celula.value}
    parcelas = []
    for numero_linha, linha in enumerate(aba.iter_rows(min_row=2, values_only=True), start=2):
        if not any(linha):
            continue
        campo = lambda nome: linha[colunas[nome]]
        try:
            parcelas.append(Parcela(
                cliente=str(campo("Cliente")).strip(),
                telefone=str(campo("Telefone") or "").strip(),
                cidade=str(campo("Cidade")).strip(),
                numero=int(campo("Parcela")),
                total=int(campo("Total de parcelas")),
                valor=para_valor(campo("Valor")),
                vencimento=para_data(campo("Vencimento")),
                situacao=str(campo("Situação") or "Em aberto"),
                pago_em=para_data(campo("Data do pagamento")),
            ))
        except (TypeError, ValueError) as erro:
            log.warning("Importacao_Diaria, linha %d ignorada: %s", numero_linha, erro)
    return parcelas


def telefone_whatsapp(telefone):
    """Devolve o número no formato 55DDDNÚMERO, ou None se estiver incompleto."""
    digitos = re.sub(r"\D", "", telefone)
    if digitos.startswith("55") and len(digitos) in (12, 13):
        digitos = digitos[2:]
    if len(digitos) == 11 and digitos[2] == "9":
        return "55" + digitos
    return None


# ---------------------------------------------------------------- textos

def reais(valor):
    return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")


def cronograma(parcelas_do_cliente):
    """Ex.: '1ª parcela dia 26/10 PAGO / 2ª parcela dia 02/11 (Pendente)'."""
    partes = []
    for p in sorted(parcelas_do_cliente, key=lambda p: p.numero):
        situacao = "PAGO" if p.paga else "(Pendente)"
        partes.append(f"{p.numero}ª parcela dia {p.vencimento:%d/%m} {situacao}")
    return " / ".join(partes)


class Redator:
    """Escreve as mensagens. Com a chave da OpenAI, a IA varia o texto;
    sem a chave, usa modelos prontos com variações, e o robô funciona igual."""

    SAUDACOES = ["Olá, {nome}!", "Oi, {nome}, tudo bem?", "{periodo}, {nome}!", "{nome}, tudo certo?"]
    COBRANCAS = [
        "Passando para lembrar que a sua {numero}ª parcela, de {valor}, vence hoje ({vencimento}).",
        "Hoje ({vencimento}) é o vencimento da sua {numero}ª parcela, no valor de {valor}.",
        "Só um lembrete: a {numero}ª parcela, de {valor}, vence hoje, {vencimento}.",
    ]
    FECHAMENTOS = [
        "Se já pagou, pode desconsiderar. Obrigado!",
        "Qualquer dúvida, é só responder esta mensagem.",
        "Se precisar de ajuda com o pagamento, me chame por aqui.",
    ]

    def __init__(self, sorteio):
        self.sorteio = sorteio
        hora = datetime.now().hour
        self.periodo = "Bom dia" if hora < 12 else "Boa tarde" if hora < 18 else "Boa noite"
        self.ia = None
        if os.environ.get("OPENAI_API_KEY"):
            try:
                from openai import OpenAI
                self.ia = OpenAI()
                log.info("Mensagens escritas com a OpenAI.")
            except ImportError:
                log.warning("OPENAI_API_KEY definida, mas o pacote openai não está instalado. Usando modelos prontos.")

    def _variar_com_ia(self, texto):
        """Pede à IA outra forma de dizer a mesma coisa, sem mudar valores e datas."""
        try:
            resposta = self.ia.chat.completions.create(
                model=os.environ.get("OPENAI_MODEL", "gpt-4o-mini"),
                messages=[
                    {"role": "system", "content": (
                        "Reescreva a mensagem de cobrança em português do Brasil, com outras palavras, "
                        "tom cordial e curto. Não mude nomes, valores, datas nem números de parcela. "
                        "Responda só com a mensagem.")},
                    {"role": "user", "content": texto},
                ],
                temperature=0.9,
            )
            novo = resposta.choices[0].message.content.strip()
            # Se a IA mexeu em algum valor ou data, fica o texto original.
            if all(numero in novo for numero in re.findall(r"\d+[\d,./]*", texto)):
                return novo
            log.warning("A IA mudou um valor ou data; mantive o texto original.")
        except Exception as erro:  # a cobrança não pode parar por causa da IA
            log.warning("OpenAI indisponível (%s); usando o modelo pronto.", erro)
        return texto

    def cobranca(self, parcela):
        nome = parcela.cliente.split()[0]
        texto = " ".join([
            self.sorteio.choice(self.SAUDACOES).format(nome=nome, periodo=self.periodo),
            self.sorteio.choice(self.COBRANCAS).format(
                numero=parcela.numero, valor=reais(parcela.valor),
                vencimento=f"{parcela.vencimento:%d/%m}"),
            self.sorteio.choice(self.FECHAMENTOS),
        ])
        return self._variar_com_ia(texto) if self.ia else texto

    def confirmacao(self, parcela, parcelas_do_cliente):
        nome = parcela.cliente.split()[0]
        return (f"{self.sorteio.choice(self.SAUDACOES).format(nome=nome, periodo=self.periodo)} "
                f"Recebemos o pagamento da sua {parcela.numero}ª parcela, obrigado!\n\n"
                f"Seu cronograma:\n{cronograma(parcelas_do_cliente).replace(' / ', chr(10))}")

    @staticmethod
    def resumo_cidade(cidade, hoje, cobrados, pagos_hoje, em_aberto):
        return (f"Resumo da cobrança de {hoje:%d/%m} — {cidade}\n"
                f"• Cobranças enviadas hoje: {cobrados}\n"
                f"• Pagamentos recebidos hoje: {pagos_hoje}\n"
                f"• Parcelas em aberto até hoje: {len(em_aberto)} ({reais(sum(p.valor for p in em_aberto))})")


# ---------------------------------------------------------------- WhatsApp

class WhatsAppSimulado:
    """Faz o papel da API de WhatsApp (Evolution API, Z-API...).

    Num trabalho real, só esta classe muda: enviar() passa a chamar a API.
    Os cuidados contra bloqueio (intervalo aleatório e limite por hora)
    ficam aqui, valendo para qualquer API.
    """

    def __init__(self, sorteio, rapido):
        self.sorteio = sorteio
        self.rapido = rapido
        self.enviadas = []          # (hora, para, telefone, texto)

    def pode_enviar(self):
        return len(self.enviadas) < LIMITE_POR_HORA

    def enviar(self, para, telefone, texto):
        if self.enviadas:
            espera = self.sorteio.uniform(*INTERVALO)
            if self.rapido:
                espera /= 40
            time.sleep(espera)
        self.enviadas.append((datetime.now(), para, telefone, texto))
        log.info("WhatsApp → %s: %s", para, texto.splitlines()[0][:70])


# ---------------------------------------------------------------- planilha

def atualizar_historico(aba, parcelas):
    """Grava cada parcela no histórico, sem duplicar. Devolve {chave: linha}."""
    colunas = {celula.value: celula.column for celula in aba[1] if celula.value}
    linhas = {}
    for linha in range(2, aba.max_row + 1):
        cliente, numero = aba.cell(linha, colunas["Cliente"]).value, aba.cell(linha, colunas["Parcela"]).value
        if cliente:
            linhas[(str(cliente).strip().lower(), int(numero))] = linha

    for p in parcelas:
        linha = linhas.get(p.chave)
        if linha is None:
            linha = aba.max_row + 1
            linhas[p.chave] = linha
        valores = {"Cliente": p.cliente, "Telefone": p.telefone, "Cidade": p.cidade,
                   "Parcela": p.numero, "Total de parcelas": p.total, "Valor": p.valor,
                   "Vencimento": p.vencimento, "Situação": "Pago" if p.paga else "Em aberto",
                   "Data do pagamento": p.pago_em}
        for nome, valor in valores.items():
            celula = aba.cell(linha, colunas[nome], valor)
            celula.font = Font(name=FONTE)
            if nome == "Valor":
                celula.number_format = '"R$" #,##0.00'
            elif nome in ("Vencimento", "Data do pagamento"):
                celula.number_format = "DD/MM/YYYY"
            elif nome in ("Parcela", "Total de parcelas", "Situação"):
                celula.alignment = Alignment(horizontal="center")
        aba.cell(linha, colunas["Situação"]).fill = VERDE if p.paga else PatternFill(fill_type=None)
    return linhas, colunas


def ja_enviado_hoje(aba, tipo, para, hoje):
    for data_hora, tipo_envio, destino, _, situacao, _ in aba.iter_rows(min_row=2, values_only=True):
        if (isinstance(data_hora, datetime) and data_hora.date() == hoje
                and tipo_envio == tipo and destino == para and situacao == "Enviada"):
            return True
    return False


def anotar_envio(aba, tipo, para, telefone, situacao, mensagem):
    aba.append([datetime.now().replace(microsecond=0), tipo, para, telefone, situacao, mensagem])
    linha = aba.max_row
    aba.cell(linha, 1).number_format = "DD/MM/YYYY HH:MM"
    aba.cell(linha, 6).alignment = Alignment(wrap_text=True, vertical="top")
    aba.cell(linha, 5).fill = {"Enviada": VERDE, "Adiada": AMARELO}.get(situacao, VERMELHO)


def planilha_livre(caminho):
    """No Windows, a planilha aberta no Excel fica travada para gravação."""
    try:
        with open(caminho, "r+b"):
            return True
    except PermissionError:
        return False


# ---------------------------------------------------------------- tela de demonstração

def gerar_html(enviadas, hoje):
    conversas = defaultdict(list)
    for hora, para, telefone, texto in enviadas:
        conversas[(para, telefone)].append((hora, texto))
    blocos = []
    for (para, telefone), mensagens in conversas.items():
        grupo = para.startswith("Grupo")
        bolhas = "".join(
            f'<div class="bolha"><p>{html.escape(texto).replace(chr(10), "<br>")}</p>'
            f'<span>{hora:%H:%M} ✓✓</span></div>' for hora, texto in mensagens)
        blocos.append(
            f'<section class="conversa"><header><b>{"👥" if grupo else "👤"} {html.escape(para)}</b>'
            f'<small>{html.escape(telefone)}</small></header>{bolhas}</section>')
    CONVERSA_HTML.write_text(f"""<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>WhatsApp simulado — {hoje:%d/%m/%Y}</title>
<style>
body {{ font-family: system-ui, sans-serif; background: #eae6df; margin: 0; padding: 16px; }}
h1 {{ font-size: 18px; color: #075e54; }}
.aviso {{ background: #fff4cc; padding: 8px 12px; border-radius: 8px; font-size: 14px; }}
.grade {{ display: grid; grid-template-columns: repeat(auto-fill, minmax(300px, 1fr)); gap: 16px; margin-top: 16px; }}
.conversa {{ background: #efeae2; border-radius: 10px; overflow: hidden; box-shadow: 0 1px 3px #0003; }}
header {{ background: #075e54; color: #fff; padding: 10px 12px; display: flex; flex-direction: column; }}
header small {{ opacity: .8; }}
.bolha {{ background: #d9fdd3; margin: 10px 10px 10px 40px; padding: 8px 10px; border-radius: 8px; font-size: 14px; }}
.bolha p {{ margin: 0 0 4px; }}
.bolha span {{ display: block; text-align: right; font-size: 11px; color: #667; }}
</style></head><body>
<h1>WhatsApp simulado — cobrança de {hoje:%d/%m/%Y}</h1>
<p class="aviso">Demonstração: nenhuma mensagem foi enviada de verdade. {len(enviadas)} mensagens.</p>
<div class="grade">{"".join(blocos)}</div>
</body></html>""", encoding="utf-8")


# ---------------------------------------------------------------- robô

def rodar(hoje, rapido, abrir):
    if not PLANILHA.exists():
        sys.exit("Não achei a planilha_cobranca.xlsx. Rode antes: python gerar_planilha.py")
    if not planilha_livre(PLANILHA):
        sys.exit("A planilha está aberta no Excel. Feche e rode de novo (nada foi enviado).")

    livro = load_workbook(PLANILHA)
    parcelas = ler_importacao(livro["Importacao_Diaria"])
    log.info("Relatório do dia: %d parcelas lidas.", len(parcelas))
    historico = livro["Historico_Parcelas"]
    envios = livro["Envios"]
    linhas, colunas = atualizar_historico(historico, parcelas)

    por_cliente = defaultdict(list)
    for p in parcelas:
        por_cliente[p.cliente].append(p)

    sorteio = random.Random()
    redator = Redator(sorteio)
    whatsapp = WhatsAppSimulado(sorteio, rapido)
    contagem = defaultdict(int)

    def mandar(tipo, para, telefone, texto):
        numero = telefone_whatsapp(telefone)
        if numero is None:
            anotar_envio(envios, tipo, para, telefone, "Erro", f"Telefone inválido ou vazio: '{telefone}'")
            log.warning("%s: telefone inválido ('%s'), não enviei.", para, telefone)
            contagem["erros"] += 1
            return False
        if ja_enviado_hoje(envios, tipo, para, hoje):
            log.info("%s: já recebeu hoje (%s), pulei.", para, tipo.lower())
            contagem["repetidas"] += 1
            return False
        if not whatsapp.pode_enviar():
            anotar_envio(envios, tipo, para, telefone, "Adiada", "Limite por hora atingido; vai na próxima rodada.")
            contagem["adiadas"] += 1
            return False
        whatsapp.enviar(para, numero, texto)
        anotar_envio(envios, tipo, para, telefone, "Enviada", texto)
        contagem[tipo] += 1
        return True

    # 1) Confirmação de pagamento, com cronograma
    coluna_confirmacao = colunas["Confirmação enviada"]
    for p in parcelas:
        celula = historico.cell(linhas[p.chave], coluna_confirmacao)
        if not p.paga or celula.value:
            continue
        if p.pago_em and p.pago_em >= hoje - timedelta(days=DIAS_PARA_CONFIRMAR):
            if mandar("Confirmação", p.cliente, p.telefone, redator.confirmacao(p, por_cliente[p.cliente])):
                celula.value = "Sim"
        else:
            celula.value = "Antiga"  # pagamento de antes do robô; não manda agora

    # 2) Cobrança de quem vence hoje
    cobrados_por_cidade = defaultdict(int)
    for p in parcelas:
        if p.vencimento == hoje and not p.paga:
            if mandar("Cobrança", p.cliente, p.telefone, redator.cobranca(p)):
                cobrados_por_cidade[p.cidade] += 1

    # 3) Resumo por cidade, nos grupos
    for cidade in sorted({p.cidade for p in parcelas}):
        da_cidade = [p for p in parcelas if p.cidade == cidade]
        pagos_hoje = sum(1 for p in da_cidade if p.paga and p.pago_em == hoje)
        em_aberto = [p for p in da_cidade if not p.paga and p.vencimento <= hoje]
        texto = Redator.resumo_cidade(cidade, hoje, cobrados_por_cidade[cidade], pagos_hoje, em_aberto)
        mandar("Resumo", f"Grupo Cobrança {cidade}", "41 99000-0000", texto)

    try:
        livro.save(PLANILHA)
    except PermissionError:
        log.error("Não consegui salvar a planilha (aberta no Excel?). Os envios estão no robo.log.")
    gerar_html(whatsapp.enviadas, hoje)

    log.info("Pronto: %d cobranças, %d confirmações, %d resumos, %d erros de telefone, %d repetidas evitadas, %d adiadas.",
             contagem["Cobrança"], contagem["Confirmação"], contagem["Resumo"],
             contagem["erros"], contagem["repetidas"], contagem["adiadas"])
    if abrir and whatsapp.enviadas:
        webbrowser.open(CONVERSA_HTML.resolve().as_uri())


def main():
    argumentos = argparse.ArgumentParser(description=__doc__.splitlines()[0])
    argumentos.add_argument("--data", help="simula outro dia (AAAA-MM-DD)")
    argumentos.add_argument("--rapido", action="store_true", help="demonstração: quase sem espera entre mensagens")
    argumentos.add_argument("--abrir", action="store_true", help="abre a tela do WhatsApp simulado no fim")
    opcoes = argumentos.parse_args()

    SAIDA.mkdir(exist_ok=True)
    logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s",
                        datefmt="%d/%m/%Y %H:%M:%S",
                        handlers=[logging.FileHandler(SAIDA / "robo.log", encoding="utf-8"),
                                  logging.StreamHandler()])
    hoje = date.fromisoformat(opcoes.data) if opcoes.data else date.today()
    rodar(hoje, opcoes.rapido, opcoes.abrir)


if __name__ == "__main__":
    main()
