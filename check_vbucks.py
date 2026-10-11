"""
Script standalone para verificar V-Bucks do Salve o Mundo (Fortnite).
Roda em loop até encontrar a atualização diária ou atingir o timeout.

Variáveis de ambiente necessárias:
  WHATSAPP_TELEFONE  ex.: 5511997590616
  WHATSAPP_APIKEY    chave fornecida pelo CallMeBot
"""
import json
import logging
import os
import re
import sys
import time
import urllib.parse
import urllib.request
from datetime import datetime, timezone

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")

URL_FONTE = "https://vbucksdaily.com/"
AGENTE = "AlertasVBucksWhatsApp/1.0 (uso pessoal; leitura diária da página inicial)"
URL_CALLMEBOT = "https://api.callmebot.com/whatsapp.php"
MISSOES_PT = {
    "evacuate the shelter": "Evacuar o Abrigo",
    "evacuate the survivors": "Evacuar os Sobreviventes",
    "retrieve the data": "Recuperar os Dados",
    "ride the lightning": "Cavalgar o Raio",
    "deliver the bomb": "Entregar a Bomba",
    "build the radar grid": "Construir a Grade de Radar",
    "destroy the encampments": "Destruir os Acampamentos",
    "fight the storm": "Enfrentar a Tempestade",
    "fight category 4 storm": "Enfrentar a Tempestade Categoria 4",
    "category 4 fight the storm": "Enfrentar a Tempestade Categoria 4",
    "repair the shelter": "Consertar o Abrigo",
}

# Timeout: 11 minutos a partir do início (cobre 20:59 até 21:10 BRT)
TIMEOUT_SEGUNDOS = 11 * 60
INTERVALO_RETRY = 30


def ler_fonte() -> dict:
    req = urllib.request.Request(
        URL_FONTE,
        headers={"User-Agent": AGENTE, "Accept": "text/html"},
    )
    with urllib.request.urlopen(req, timeout=30) as resp:
        html = resp.read(3_000_000).decode("utf-8", errors="replace")
    achou = re.search(
        r'<script type="application/json" id="mission-bootstrap">(.*?)</script>',
        html,
        flags=re.S,
    )
    if not achou:
        raise RuntimeError("Formato inesperado: mission-bootstrap não encontrado.")
    return json.loads(achou.group(1)).get("snapshot") or {}


def alertas_de_vbucks(snapshot: dict) -> list:
    alertas = []
    for m in snapshot.get("missions") or []:
        quantidade = int(m.get("quantity") or 0)
        if quantidade > 0 and (
            m.get("category") == "vbucks"
            or "v-bucks" in str(m.get("reward") or "").lower()
        ):
            alertas.append(m)
    return alertas


def montar_mensagem(alertas: list) -> str:
    total = sum(int(a.get("quantity") or 0) for a in alertas)
    linhas = [f"*{total} V-Bucks disponíveis no Salve o Mundo!*", ""]
    for a in alertas:
        tipo = a.get("type") or "Missão"
        nome = MISSOES_PT.get(tipo.lower(), tipo)
        detalhe = " - ".join(filter(None, [nome, a.get("biome")]))
        linhas.append(
            f"- {a.get('zone')} (poder {a.get('pl')}): {a.get('quantity')} V-Bucks\n  {detalhe}"
        )
    return "\n".join(linhas)


def enviar(texto: str) -> None:
    telefone = "".join(c for c in os.environ.get("WHATSAPP_TELEFONE", "") if c.isdigit())
    apikey = os.environ.get("WHATSAPP_APIKEY", "")
    if not telefone or not apikey:
        raise RuntimeError("Defina WHATSAPP_TELEFONE e WHATSAPP_APIKEY.")
    consulta = urllib.parse.urlencode({"phone": "+" + telefone, "text": texto, "apikey": apikey})
    with urllib.request.urlopen(f"{URL_CALLMEBOT}?{consulta}", timeout=60) as resp:
        corpo = resp.read().decode("utf-8", errors="replace")
        status = resp.status
    if status != 200 or "error" in corpo.lower():
        raise RuntimeError(f"CallMeBot recusou o envio ({status}): {corpo[:200]}")


def main() -> None:
    inicio = time.monotonic()
    tentativa = 0

    while True:
        tentativa += 1
        decorrido = time.monotonic() - inicio

        if decorrido >= TIMEOUT_SEGUNDOS:
            logging.warning("Timeout atingido. Enviando aviso de ausência de V-Bucks.")
            try:
                hoje_fmt = datetime.now(timezone.utc).strftime("%d/%m/%Y")
                enviar(f"Hoje ({hoje_fmt}) não há V-Bucks disponíveis no Salve o Mundo.")
            except Exception as erro:
                logging.error("Falha ao enviar aviso: %r", erro)
            sys.exit(0)

        # Recalcula "hoje" a cada iteração para cruzar a meia-noite UTC corretamente
        hoje = datetime.now(timezone.utc).strftime("%Y-%m-%d")

        try:
            snapshot = ler_fonte()
            rotation_id = snapshot.get("rotationId", "")
            logging.info(
                "Tentativa %d | rotationId=%s | esperado=%s | %.0fs decorridos",
                tentativa, rotation_id, hoje, decorrido,
            )

            if rotation_id == hoje:
                alertas = alertas_de_vbucks(snapshot)
                if alertas:
                    enviar(montar_mensagem(alertas))
                    logging.info("Sucesso: %d alerta(s) enviado(s) via WhatsApp.", len(alertas))
                else:
                    hoje_fmt = datetime.now(timezone.utc).strftime("%d/%m/%Y")
                    enviar(f"Hoje ({hoje_fmt}) não há V-Bucks disponíveis no Salve o Mundo.")
                    logging.info("Sem V-Bucks hoje; aviso enviado.")
                sys.exit(0)

        except Exception as erro:
            logging.warning("Tentativa %d falhou: %r", tentativa, erro)

        restante = TIMEOUT_SEGUNDOS - (time.monotonic() - inicio)
        if restante <= 0:
            continue
        time.sleep(min(INTERVALO_RETRY, restante))


if __name__ == "__main__":
    main()
