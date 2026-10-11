"""
Script standalone para verificar V-Bucks do Salve o Mundo (Fortnite).
Roda em loop até encontrar a atualização diária ou atingir o timeout.
Usa Supabase para garantir que o alerta seja enviado apenas uma vez por dia.

Variáveis de ambiente necessárias:
  WHATSAPP_TELEFONE  ex.: 5511997590616
  WHATSAPP_APIKEY    chave fornecida pelo CallMeBot
  SUPABASE_URL       ex.: https://xyzxyz.supabase.co
  SUPABASE_KEY       anon/service_role key do projeto
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


# ---------------------------------------------------------------------------
# Supabase (REST puro, sem dependências externas)
# ---------------------------------------------------------------------------

def _supabase_headers() -> dict:
    key = os.environ.get("SUPABASE_SERVICE_ROLE_KEY") or os.environ.get("SUPABASE_KEY", "")
    return {
        "apikey": key,
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
        "Prefer": "return=minimal",
    }


def alerta_ja_enviado(data: str) -> bool:
    """Retorna True se já existe registro para essa data no Supabase."""
    url = os.environ.get("SUPABASE_URL", "").rstrip("/")
    if not url:
        logging.warning("SUPABASE_URL não definida; checagem de duplicata ignorada.")
        return False
    endpoint = f"{url}/rest/v1/alertas_vbucks?data=eq.{data}&select=data"
    req = urllib.request.Request(endpoint, headers=_supabase_headers())
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            registros = json.loads(resp.read().decode())
            return len(registros) > 0
    except Exception as erro:
        logging.warning("Supabase: falha na leitura (%r); prosseguindo sem checagem.", erro)
        return False


def registrar_envio(data: str, tipo: str, quantidade: int | None = None) -> None:
    """Insere ou atualiza o registro de envio do dia no Supabase."""
    url = os.environ.get("SUPABASE_URL", "").rstrip("/")
    if not url:
        return
    endpoint = f"{url}/rest/v1/alertas_vbucks"
    payload = {"data": data, "tipo": tipo}
    if quantidade is not None:
        payload["quantidade"] = quantidade
    headers = _supabase_headers()
    # upsert: se já existir o mesmo 'data', atualiza
    headers["Prefer"] = "resolution=merge-duplicates,return=minimal"
    dados = json.dumps(payload).encode()
    req = urllib.request.Request(endpoint, data=dados, headers=headers, method="POST")
    try:
        with urllib.request.urlopen(req, timeout=10):
            pass
        logging.info("Supabase: registro salvo (%s | %s).", data, tipo)
    except Exception as erro:
        logging.warning("Supabase: falha ao registrar envio (%r).", erro)


# ---------------------------------------------------------------------------
# Lógica de V-Bucks
# ---------------------------------------------------------------------------

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


def enviar_whatsapp(texto: str) -> None:
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


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main() -> None:
    inicio = time.monotonic()
    tentativa = 0

    while True:
        tentativa += 1
        decorrido = time.monotonic() - inicio

        # Recalcula "hoje" a cada iteração para cruzar a meia-noite UTC corretamente
        hoje = datetime.now(timezone.utc).strftime("%Y-%m-%d")
        hoje_fmt = datetime.now(timezone.utc).strftime("%d/%m/%Y")

        if decorrido >= TIMEOUT_SEGUNDOS:
            logging.warning("Timeout atingido.")
            if alerta_ja_enviado(hoje):
                logging.info("Alerta de hoje já foi enviado anteriormente; encerrando sem reenvio.")
                sys.exit(0)
            try:
                enviar_whatsapp(f"Hoje ({hoje_fmt}) não há V-Bucks disponíveis no Salve o Mundo.")
                registrar_envio(hoje, "sem_vbucks")
            except Exception as erro:
                logging.error("Falha ao enviar aviso de ausência: %r", erro)
            sys.exit(0)

        try:
            snapshot = ler_fonte()
            rotation_id = snapshot.get("rotationId", "")
            logging.info(
                "Tentativa %d | rotationId=%s | esperado=%s | %.0fs decorridos",
                tentativa, rotation_id, hoje, decorrido,
            )

            if rotation_id == hoje:
                if alerta_ja_enviado(hoje):
                    logging.info("Alerta de hoje já foi enviado anteriormente; encerrando sem reenvio.")
                    sys.exit(0)

                alertas = alertas_de_vbucks(snapshot)
                if alertas:
                    total = sum(int(a.get("quantity") or 0) for a in alertas)
                    enviar_whatsapp(montar_mensagem(alertas))
                    registrar_envio(hoje, "vbucks", total)
                    logging.info("Sucesso: %d alerta(s) enviado(s) via WhatsApp.", len(alertas))
                else:
                    enviar_whatsapp(f"Hoje ({hoje_fmt}) não há V-Bucks disponíveis no Salve o Mundo.")
                    registrar_envio(hoje, "sem_vbucks")
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
