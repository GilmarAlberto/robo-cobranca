#!/usr/bin/env bash
# Rodar o robô no Linux: dois cliques (ou ./rodar_robo.sh no terminal).
# Na primeira vez prepara tudo sozinho; nas próximas, só roda.
cd "$(dirname "$0")" || exit 1

if ! command -v python3 >/dev/null 2>&1; then
    echo "Python 3 não encontrado. Instale pelo gerenciador de programas do seu Linux."
    read -rp "Aperte Enter para fechar."
    exit 1
fi

if [ ! -d .venv ]; then
    echo "Primeira vez: preparando o robô (leva alguns minutos)..."
    python3 -m venv .venv &&
    .venv/bin/pip install -q -r requirements.txt || {
        echo "Não consegui preparar o robô. Confira a internet e tente de novo."
        rm -rf .venv
        read -rp "Aperte Enter para fechar."
        exit 1
    }
fi

[ -f planilha_cobranca.xlsx ] || .venv/bin/python gerar_planilha.py
.venv/bin/python robo_cobranca.py --rapido --abrir "$@"
read -rp "Pronto. Aperte Enter para fechar."
