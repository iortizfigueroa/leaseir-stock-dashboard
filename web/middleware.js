// Edge Middleware (Vercel): protege todo el dashboard con contraseña.
// La contraseña se configura en Vercel → Settings → Environment Variables → DASHBOARD_PASSWORD.
// Si no hay variable configurada, no bloquea (para no dejarte fuera por accidente).

export const config = {
  matcher: "/((?!login\\.html|api/login|favicon\\.ico|_vercel).*)",
};

async function sha256hex(s) {
  const buf = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(s));
  return Array.from(new Uint8Array(buf)).map(b => b.toString(16).padStart(2, "0")).join("");
}

export default async function middleware(request) {
  const expected = process.env.DASHBOARD_PASSWORD || "";
  if (!expected) return; // sin contraseña configurada → acceso libre

  const cookie = request.headers.get("cookie") || "";
  const m = cookie.match(/(?:^|;\s*)lsr_auth=([0-9a-f]{64})/);
  if (m && m[1] === (await sha256hex(expected))) return; // autenticado → continuar

  return Response.redirect(new URL("/login.html", request.url), 302);
}
