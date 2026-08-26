// API del plan de Previsión — guarda el plan completo (salidas/mes, mix, compras)
// en data/master/prevision.json del propio repo. Compartido para todo el equipo.
// GET  → {plan: {...}}
// POST {plan: {...}} → sustituye el plan y commitea a main.
// Protegida por el middleware (cookie). Usa el mismo env GITHUB_TOKEN que los mínimos.

const REPO = "iortizfigueroa/leaseir-stock-dashboard";
const FILE = "data/master/prevision.json";

export default async function handler(req, res) {
  const token = process.env.GITHUB_TOKEN;
  const gh = (path, opts = {}) =>
    fetch(`https://api.github.com${path}`, {
      ...opts,
      headers: {
        Authorization: `Bearer ${token}`,
        Accept: "application/vnd.github+json",
        "User-Agent": "leaseir-dashboard",
        ...(opts.headers || {}),
      },
    });

  if (req.method === "GET") {
    try {
      let data = {};
      if (token) {
        // Contents API con token: SIN cache de CDN (raw.githubusercontent cachea ~5 min)
        const r = await gh(`/repos/${REPO}/contents/${FILE}?ref=main`);
        if (r.ok) {
          const j = await r.json();
          try { data = JSON.parse(Buffer.from(j.content, "base64").toString("utf-8")); } catch (e) {}
        }
      } else {
        const r = await fetch(
          `https://raw.githubusercontent.com/${REPO}/main/${FILE}`,
          { headers: { "User-Agent": "leaseir-dashboard" }, cache: "no-store" }
        );
        data = r.ok ? await r.json() : {};
      }
      res.setHeader("Cache-Control", "no-store");
      res.status(200).json({ plan: data });
    } catch (e) {
      res.status(200).json({ plan: {} });
    }
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }
  if (!token) {
    res.status(500).json({ error: "GITHUB_TOKEN no configurado en Vercel" });
    return;
  }
  const plan = (req.body || {}).plan;
  if (plan == null || typeof plan !== "object" || Array.isArray(plan)) {
    res.status(400).json({ error: "plan inválido" });
    return;
  }
  const content = JSON.stringify(plan, null, 1);
  if (content.length > 200000) {
    res.status(400).json({ error: "plan demasiado grande" });
    return;
  }

  for (let attempt = 0; attempt < 2; attempt++) {
    const cur = await gh(`/repos/${REPO}/contents/${FILE}?ref=main`);
    let sha = null;
    if (cur.ok) sha = (await cur.json()).sha;

    const body = {
      message: "Plan de prevision actualizado (via dashboard)",
      content: Buffer.from(content).toString("base64"),
      branch: "main",
    };
    if (sha) body.sha = sha;

    const put = await gh(`/repos/${REPO}/contents/${FILE}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (put.ok) {
      res.status(200).json({ ok: true });
      return;
    }
    if (put.status !== 409) {
      const t = await put.text();
      res.status(502).json({ error: `GitHub ${put.status}: ${t.substring(0, 200)}` });
      return;
    }
  }
  res.status(409).json({ error: "Conflicto al guardar; inténtalo de nuevo" });
}
