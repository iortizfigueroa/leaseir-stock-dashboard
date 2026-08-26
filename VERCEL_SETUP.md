# Dashboard v2 en Vercel — instrucciones de puesta en marcha

El dashboard de GitHub Pages sigue funcionando igual que siempre. Esta versión v2 es
paralela: se genera con los mismos datos en cada corrida diaria, pero se sirve desde
Vercel con contraseña y trae dos pestañas nuevas (🔔 Alertas y 📈 Previsión).

## Paso 1 — Subir los archivos al repo (GitHub web)

Del zip, sube cada carpeta a su sitio en el repo (Add file → Upload files):

- `src/` → los 4 archivos (2 .py y 2 .html) dentro de la carpeta `src` del repo
- `scripts/build.sh` → dentro de la carpeta `scripts`
- `web/` → arrastra la carpeta `web` entera en la RAÍZ del repo (crea la carpeta nueva
  con `index.html`, `login.html`, `middleware.js` y `api/login.js`)
- `VERCEL_SETUP.md` → opcional, en la raíz (es este documento)

## Paso 2 — Un cambio de una línea en el workflow

Abre `.github/workflows/update-daily.yml` en GitHub, pulsa el lápiz (editar) y cambia
la línea:

    git add data/ejercicios/ docs/

por:

    git add data/ejercicios/ docs/ web/

Guarda (Commit changes). Sin esto, la versión Vercel no se refrescaría cada día.

## Paso 3 — Crear el proyecto en Vercel (una sola vez)

1. Entra en vercel.com → **Add New… → Project**
2. **Import** el repositorio `iortizfigueroa/leaseir-stock-dashboard`
   (la primera vez te pedirá conectar tu cuenta de GitHub)
3. En la pantalla de configuración, ANTES de darle a Deploy:
   - **Root Directory**: pulsa Edit y elige `web`
   - **Framework Preset**: `Other`
   - **Environment Variables**: añade una variable
     - Name: `DASHBOARD_PASSWORD`
     - Value: la contraseña que quieras para el equipo
4. **Deploy**

Vercel te dará una URL tipo `https://leaseir-stock-dashboard.vercel.app`.
Al entrar pedirá la contraseña; una vez metida, no la vuelve a pedir en 30 días
en ese navegador. A partir de aquí, cada commit diario del workflow redespliega
la versión nueva solo (sin tocar nada).

Para cambiar la contraseña: Vercel → Settings → Environment Variables → editar
`DASHBOARD_PASSWORD` → Redeploy. (Todos tendrán que volver a loguearse.)

## Las pestañas nuevas de la v2

**🔔 Alertas** — todos los SPECs que tienen stock mínimo definido (en
`Inventario_Leaseir_Live.xlsx`), ordenados de peor a mejor cobertura: los que están
bajo mínimo en rojo, los que están a menos de 1,5× en ámbar, y una cifra de compra
sugerida para cubrir todos los mínimos. El número del badge de la pestaña es cuántos
están bajo mínimo hoy.

**📈 Previsión** — escribes qué vas a fabricar (pieza/equipo + cantidad, puedes añadir
varias líneas) y calcula, con las recetas, qué componentes consumirás y si te llega el
stock del almacén sin romper mínimos; lo que falte, con su compra estimada en euros.
El escenario se guarda en el navegador (cada uno puede tener el suyo).

## Nota honesta sobre privacidad

El repositorio de GitHub es público, y ahí es donde viven los datos (también los del
dashboard de Pages de siempre). La contraseña protege la URL de Vercel — que es la que
compartirás — pero quien conozca el repo puede ver los datos en GitHub, igual que hoy.
Si algún día quieres privacidad de verdad: se hace el repo privado (Vercel seguiría
funcionando gratis; el GitHub Pages gratuito dejaría de servir la web) y nos quedamos
solo con Vercel. Se puede hacer en cualquier momento.
