// Monitora vbucksdaily.com a cada minuto (acionado pelo pg_cron do Supabase).
// Quando a rotação do dia aparece, ou as missões de V-Bucks mudam, registra no banco
// e envia WhatsApp para todos os usuários cadastrados.

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const URL_FONTE = "https://vbucksdaily.com/";
const AGENTE = "AlertasVBucksWhatsApp/1.0 (uso pessoal; leitura da página inicial)";

const MISSOES_PT: Record<string, string> = {
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
};

// deno-lint-ignore no-explicit-any
type Json = any;
interface Missao { tipo: string; zona: string; poder: number; quantidade: number; biome: string }

async function sb(path: string, init: RequestInit = {}): Promise<Json> {
  const r = await fetch(`${SUPABASE_URL}/rest/v1/${path}`, {
    ...init,
    headers: {
      apikey: KEY,
      Authorization: `Bearer ${KEY}`,
      "Content-Type": "application/json",
      ...(init.headers ?? {}),
    },
  });
  if (!r.ok) throw new Error(`Supabase ${r.status}: ${await r.text()}`);
  const t = await r.text();
  return t ? JSON.parse(t) : null;
}

async function lerFonte(): Promise<Json> {
  const r = await fetch(URL_FONTE, { headers: { "User-Agent": AGENTE, Accept: "text/html" } });
  if (!r.ok) throw new Error(`Fonte ${r.status}`);
  const html = await r.text();
  const m = html.match(/<script type="application\/json" id="mission-bootstrap">([\s\S]*?)<\/script>/);
  if (!m) throw new Error("Formato inesperado: mission-bootstrap não encontrado.");
  return JSON.parse(m[1]).snapshot ?? {};
}

function alertasDeVbucks(snapshot: Json): Json[] {
  return (snapshot.missions ?? []).filter((m: Json) => {
    const q = Number(m.quantity ?? 0);
    return q > 0 && (m.category === "vbucks" || String(m.reward ?? "").toLowerCase().includes("v-bucks"));
  });
}

function assinatura(missoes: Missao[]): string {
  return missoes.map((m) => `${m.zona}|${m.poder}|${m.quantidade}|${m.tipo}`).sort().join(";");
}

function montarMensagem(missoes: Missao[], total: number): string {
  const linhas = [`*${total} V-Bucks disponíveis no Salve o Mundo!*`, ""];
  for (const m of missoes) {
    const detalhe = [m.tipo, m.biome].filter(Boolean).join(" - ");
    linhas.push(`- ${m.zona} (poder ${m.poder}): ${m.quantidade} V-Bucks\n  ${detalhe}`);
  }
  return linhas.join("\n");
}

async function enviarWhatsapp(texto: string, telefone: string, apikey: string) {
  const num = telefone.replace(/\D/g, "");
  if (!num || !apikey) throw new Error("Telefone ou API key inválidos.");
  const q = new URLSearchParams({ phone: "+" + num, text: texto, apikey });
  const r = await fetch(`https://api.callmebot.com/whatsapp.php?${q}`);
  const corpo = await r.text();
  if (r.status !== 200 || corpo.toLowerCase().includes("error")) {
    throw new Error(`CallMeBot recusou (${r.status}): ${corpo.slice(0, 120)}`);
  }
}

async function enviarParaTodos(texto: string): Promise<{ enviados: number; falhas: number }> {
  const usuarios: Json[] = await sb(
    "usuarios?select=telefone,whatsapp_apikey&telefone=not.is.null&whatsapp_apikey=not.is.null",
  );
  const res = await Promise.allSettled(
    usuarios.map((u) => enviarWhatsapp(texto, u.telefone, u.whatsapp_apikey)),
  );
  res.forEach((r) => r.status === "rejected" && console.warn("Falha WhatsApp:", String(r.reason)));
  const enviados = res.filter((r) => r.status === "fulfilled").length;
  return { enviados, falhas: res.length - enviados };
}

async function executar(): Promise<Json> {
  const snap = await lerFonte();
  const hoje = new Date().toISOString().slice(0, 10);
  if (snap.rotationId !== hoje) return { status: "aguardando", rotationId: snap.rotationId };

  const alertas = alertasDeVbucks(snap);
  const missoes: Missao[] = alertas.map((a) => {
    const t = String(a.type ?? "");
    return {
      tipo: MISSOES_PT[t.toLowerCase()] ?? (t || "Missão"),
      zona: a.zone,
      poder: a.pl,
      quantidade: Number(a.quantity ?? 0),
      biome: a.biome,
    };
  });
  const total = missoes.reduce((s, m) => s + m.quantidade, 0);
  const tipo = missoes.length ? "vbucks" : "sem_vbucks";
  const corpo = {
    tipo,
    quantidade: missoes.length ? total : null,
    missoes_json: missoes.length ? missoes : null,
  };

  const rows: Json[] = await sb(`alertas_vbucks?data=eq.${hoje}&select=data,tipo,missoes_json,enviado_em`);
  let atualizacao = false;

  if (!rows.length) {
    const claimed = await sb("alertas_vbucks", {
      method: "POST",
      headers: { Prefer: "resolution=ignore-duplicates,return=representation" },
      body: JSON.stringify({ data: hoje, ...corpo }),
    });
    if (!claimed?.length) return { status: "ja_processado" };
  } else {
    const row = rows[0];
    if (row.tipo === "vbucks" && !row.missoes_json && tipo === "vbucks") {
      // Registro antigo sem detalhes: completa em silêncio, sem reenviar.
      await sb(`alertas_vbucks?data=eq.${hoje}`, { method: "PATCH", body: JSON.stringify(corpo) });
      return { status: "detalhes_completados" };
    }
    if (row.tipo === tipo && assinatura(row.missoes_json ?? []) === assinatura(missoes)) {
      return { status: "sem_mudancas" };
    }
    const upd = await sb(
      `alertas_vbucks?data=eq.${hoje}&enviado_em=eq.${encodeURIComponent(row.enviado_em)}`,
      {
        method: "PATCH",
        headers: { Prefer: "return=representation" },
        body: JSON.stringify({ ...corpo, enviado_em: new Date().toISOString() }),
      },
    );
    if (!upd?.length) return { status: "ja_processado" };
    atualizacao = true;
  }

  const [ano, mes, dia] = hoje.split("-");
  let texto = missoes.length
    ? montarMensagem(missoes, total)
    : `Hoje (${dia}/${mes}/${ano}) não há V-Bucks disponíveis no Salve o Mundo.`;
  if (atualizacao) texto = "*Atualização das missões*\n" + texto;

  const envio = await enviarParaTodos(texto);
  return { status: atualizacao ? "atualizado" : "novo", tipo, total, ...envio };
}

Deno.serve(async () => {
  try {
    const out = await executar();
    console.log(JSON.stringify(out));
    return Response.json(out);
  } catch (e) {
    console.error(e);
    return Response.json({ erro: String(e) }, { status: 500 });
  }
});
