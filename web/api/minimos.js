// API de mínimos por SPEC — guarda en data/master/minimos.json del propio repo.
// GET  → {minimos: {...}}  (estado fresco, para overlay al cargar la página)
// POST {spec, minimo|null} → actualiza/borra la clave y commitea a main.
// Protegida por el middleware (cookie de sesión). Necesita env GITHUB_TOKEN
// (fine-grained, solo este repo, permiso Contents: Read & write).

const REPO = "iortizfigueroa/leaseir-stock-dashboard";
const FILE = "data/master/minimos.json";

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
      res.status(200).json({ minimos: data });
    } catch (e) {
      res.status(200).json({ minimos: {} });
    }
    return;
  }

  if (req.method !== "POST") {
    res.status(405).json({ error: "Method Not Allowed" });
    return;
  }
  if (!token) {
    res.status(500).json({ error: "GITHUB_TOKEN no configurado en Vercel (Settings → Environment Variables)" });
    return;
  }
  const { spec, minimo } = req.body || {};
  if (!spec || typeof spec !== "string" || !/^[A-Za-z0-9_\/.\-]{1,60}$/.test(spec)) {
    res.status(400).json({ error: "spec inválido" });
    return;
  }
  const val = minimo == null || minimo === "" ? null : Number(minimo);
  if (val != null && (!isFinite(val) || val < 0)) {
    res.status(400).json({ error: "mínimo inválido" });
    return;
  }

  // Hasta 2 intentos por si otro usuario guardó a la vez (conflicto de sha)
  for (let attempt = 0; attempt < 2; attempt++) {
    const cur = await gh(`/repos/${REPO}/contents/${FILE}?ref=main`);
    let sha = null;
    let data = {};
    if (cur.ok) {
      const cj = await cur.json();
      sha = cj.sha;
      try {
        data = JSON.parse(Buffer.from(cj.content, "base64").toString("utf-8"));
      } catch (e) {}
    }
    if (val == null) delete data[spec];
    else data[spec] = val;

    const body = {
      message: `Minimo ${spec} ${val == null ? "quitado" : "= " + val} (via dashboard)`,
      content: Buffer.from(
        JSON.stringify(data, Object.keys(data).sort(), 1)
      ).toString("base64"),
      branch: "main",
    };
    if (sha) body.sha = sha;

    const put = await gh(`/repos/${REPO}/contents/${FILE}`, {
      method: "PUT",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    if (put.ok) {
      res.status(200).json({ ok: true, spec: spec, minimo: val });
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
