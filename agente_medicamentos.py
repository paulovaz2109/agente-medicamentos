# -*- coding: utf-8 -*-
"""
Agente de Monitoramento de Estoque de Medicamentos
Envia alerta no WhatsApp quando medicamento vai acabar em <= 5 dias
"""

import os
from datetime import date, datetime
from openpyxl import load_workbook
from twilio.rest import Client
from dotenv import load_dotenv

# -- Carregar credenciais do arquivo .env
load_dotenv()

ACCOUNT_SID   = os.getenv("TWILIO_ACCOUNT_SID")
AUTH_TOKEN    = os.getenv("TWILIO_AUTH_TOKEN")
FROM_NUMBER   = os.getenv("TWILIO_WHATSAPP_FROM")
DESTINATARIOS = os.getenv("MEUS_NUMEROS", "").split(",")

# -- Configuracoes do agente
CAMINHO_PLANILHA = "estoque.xlsx"
DIAS_ALERTA      = 5
ABA_ESTOQUE      = "Estoque"
COLUNA_PACIENTE  = 2
COLUNA_MED       = 3
COLUNA_DURACAO   = 5


def ler_planilha():
    """Le a planilha e retorna lista de medicamentos proximos do fim."""
    wb = load_workbook(CAMINHO_PLANILHA, data_only=True)
    ws = wb[ABA_ESTOQUE]
    hoje = date.today()
    alertas = []

    for row in ws.iter_rows(min_row=3, values_only=True):
        paciente = row[COLUNA_PACIENTE - 1]
        med      = row[COLUNA_MED - 1]
        duracao  = row[COLUNA_DURACAO - 1]

        # Ignorar linhas sem dados
        if not med or not duracao:
            continue
        if isinstance(duracao, str):
            continue  # ignora "Qtd nao informada" etc.

        # Converter para date (pode vir como datetime do Excel)
        if isinstance(duracao, datetime):
            duracao = duracao.date()

        if not isinstance(duracao, date):
            continue

        dias_restantes = (duracao - hoje).days

        if dias_restantes <= DIAS_ALERTA:
            alertas.append({
                "paciente":       str(paciente) if paciente else "?",
                "medicamento":    str(med),
                "duracao":        duracao.strftime("%d/%m/%Y"),
                "dias_restantes": dias_restantes,
            })

    # Ordenar do mais urgente para o menos urgente
    alertas.sort(key=lambda x: x["dias_restantes"])
    return alertas


def formatar_mensagem(alertas):
    """Monta o texto da mensagem WhatsApp."""
    hoje_str = date.today().strftime("%d/%m/%Y")
    linhas = [
        "ALERTA DE ESTOQUE - " + hoje_str,
        "Medicamentos que vencem em ate " + str(DIAS_ALERTA) + " dias:\n",
    ]

    for a in alertas:
        d = a["dias_restantes"]
        if d < 0:
            status = "VENCIDO"
        elif d == 0:
            status = "ACABA HOJE"
        elif d == 1:
            status = "1 dia restante"
        else:
            status = str(d) + " dias restantes"

        linhas.append(
            "* " + a["paciente"] + " - " + a["medicamento"] + "\n"
            "  Dura ate: " + a["duracao"] + " (" + status + ")"
        )

    linhas.append("\n_Agente automatico - Estoque Medicamentos_")
    return "\n".join(linhas)


def enviar_whatsapp(mensagem):
    """Envia a mensagem para todos os destinatarios configurados."""
    client = Client(ACCOUNT_SID, AUTH_TOKEN)
    enviados = 0

    for destino in DESTINATARIOS:
        destino = destino.strip()
        if not destino:
            continue
        try:
            msg = client.messages.create(
                body=mensagem,
                from_=FROM_NUMBER,
                to=destino
            )
            print("OK - Enviado para " + destino + " - SID: " + msg.sid)
            enviados += 1
        except Exception as e:
            print("ERRO ao enviar para " + destino + ": " + str(e))

    return enviados


def main():
    agora = datetime.now().strftime("%Y-%m-%d %H:%M")
    print("[" + agora + "] Iniciando verificacao de estoque...")

    alertas = ler_planilha()

    if not alertas:
        print("OK - Nenhum medicamento critico. Nenhuma mensagem enviada.")
        return

    print("ATENCAO - " + str(len(alertas)) + " medicamento(s) critico(s) encontrado(s):")
    for a in alertas:
        print("   -> " + a["paciente"] + " | " + a["medicamento"] + " | " + str(a["dias_restantes"]) + " dias")

    mensagem = formatar_mensagem(alertas)
    print("\nMensagem a enviar:\n" + mensagem)

    enviados = enviar_whatsapp(mensagem)
    print("\nConcluido. " + str(enviados) + " mensagem(ns) enviada(s).")


if __name__ == "__main__":
    main()