// Vercel Serverless Function: valida la contraseña y deja una cookie de sesión (30 días).
import crypto from "node:crypto";

export default function handler(req, res) {
  if (req.method !== "POST") {
    res.statusCode = 405;
    res.end("Method Not Allowed");
    return;
  }
  const expected = process.env.DASHBOARD_PASSWORD || "";
  const pass = (req.body && req.body.password) || "";
  if (!expected || pass !== expected) {
    res.statusCode = 302;
    res.setHeader("Location", "/login.html?e=1");
    res.end();
    return;
  }
  const tok = crypto.createHash("sha256").update(expected).digest("hex");
  res.setHeader(
    "Set-Cookie",
    `lsr_auth=${tok}; Path=/; HttpOnly; Secure; SameSite=Lax; Max-Age=2592000`
  );
  res.statusCode = 302;
  res.setHeader("Location", "/");
  res.end();
}
