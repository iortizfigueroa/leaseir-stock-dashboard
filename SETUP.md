# Setup paso a paso — leaseir-stock-dashboard

Sigue este orden EXACTO. La instalación tarda ~30 minutos en total.

---

## 1) Crear el repo en GitHub

1. En github.com → **New repository**
2. Nombre: `leaseir-stock-dashboard`
3. Visibilidad: **Public** (necesario para GitHub Pages gratis)
4. NO marques "Add README" (ya viene en el zip)
5. Click **Create repository**
6. Sigue la opción "push an existing repository":
   ```bash
   cd /ruta/local/a/leaseir-stock-dashboard
   git init
   git add .
   git commit -m "Initial commit"
   git branch -M main
   git remote add origin git@github.com:iortizfigueroa/leaseir-stock-dashboard.git
   git push -u origin main
   ```

---

## 2) Configurar GitHub Pages

1. En el repo → **Settings → Pages**
2. Source: **Deploy from a branch**
3. Branch: **main**, folder: **/docs**
4. Save
5. Espera ~1 minuto → la URL será:
   `https://iortizfigueroa.github.io/leaseir-stock-dashboard/`

---

## 3) Crear Service Account para Google Drive

(Esto es lo más liado pero solo se hace una vez)

### 3.1 Crear proyecto en Google Cloud

1. Entra en [console.cloud.google.com](https://console.cloud.google.com)
2. Arriba a la izquierda, junto al logo, click el dropdown → **New project**
3. Nombre: `leaseir-dashboard-gha`, click **Create**

### 3.2 Habilitar Drive API

1. En el buscador de arriba escribe **Google Drive API** → entra
2. Click **Enable**

### 3.3 Crear Service Account

1. **IAM & Admin → Service Accounts → Create service account**
2. Nombre: `gha-leaseir-dashboard`
3. Click **Create and continue** → **Done** (sin asignar roles)
4. En la lista de Service Accounts, click el que acabas de crear
5. Pestaña **Keys → Add key → Create new key → JSON**
6. Se descargará un fichero `.json`. **GUARDA ESTE FICHERO BIEN.**

### 3.4 Compartir la carpeta de Drive con el Service Account

1. Abre el JSON descargado → busca el campo `client_email`
   (algo como `gha-leaseir-dashboard@leaseir-dashboard-gha.iam.gserviceaccount.com`)
2. Ve a Google Drive → click derecho en la carpeta donde están los `ejercicio *.xlsx`
3. **Compartir → añade el client_email** como **Lector**
4. Copia el **ID de la carpeta**: en la URL de Drive `https://drive.google.com/drive/folders/AbCdEfGhIjKlMnOpQrSt` → el ID es `AbCdEfGhIjKlMnOpQrSt`

---

## 4) Configurar Secrets en GitHub

En tu repo → **Settings → Secrets and variables → Actions → New repository secret**

Crea dos secrets:

| Nombre | Valor |
|---|---|
| `GDRIVE_SA_KEY` | El **contenido completo** del fichero JSON del Service Account (ábrelo con notepad y copia/pega TODO) |
| `GDRIVE_FOLDER_ID` | El ID de la carpeta de Drive (del paso 3.4) |

---

## 5) Probar manualmente

1. En el repo → **Actions → Update dashboard daily → Run workflow → Run workflow** (botón verde)
2. Espera ~2 minutos
3. Si va bien:
   - El job aparece en verde ✓
   - Verás un commit nuevo en `main` con el HTML actualizado
   - La URL https://iortizfigueroa.github.io/leaseir-stock-dashboard/ ya muestra el dashboard

Si falla:
- Click en el job rojo → mira los logs del paso que falló
- Errores típicos:
  - `403 Forbidden` → la carpeta de Drive no está compartida con el Service Account
  - `GDRIVE_*` no configurado → los secrets están mal escritos
  - Python error → revisa logs y dime

---

## 6) Confirmar el cron automático

Una vez probado manualmente, el cron se ejecutará solo cada día a las 17:00 hora Madrid.

El workflow lanza DOS corridas (15:00 UTC y 16:00 UTC) para cubrir verano/invierno. La segunda no hace nada si no hay cambios.

---

## Mantenimiento futuro

- **Añadir un ejercicio nuevo cada día**: simplemente sube el `ejercicio DD-MM-YYYY.xlsx` a la carpeta de Drive. El workflow lo coge solo a las 17:00.
- **Cambiar el código del dashboard**: edita `src/*.py` o `src/vision_html_template_full.html`, commitea y pushea. La siguiente corrida usará el nuevo código.
- **Actualizar el master (escandallos, Live, etc.)**: edita en `data/master/`, commitea y pushea.
- **Histórico**: el repo guarda commits diarios → puedes volver atrás a cualquier día.
- **Forzar actualización ahora**: Actions → Update dashboard daily → Run workflow.

---

## Estructura del repo

```
leaseir-stock-dashboard/
├── .github/workflows/
│   └── update-daily.yml         # Cron + steps
├── src/                          # Código Python + template HTML
│   ├── vision_html_full.py
│   ├── vision_html_template_full.html
│   └── ... otros .py
├── data/
│   ├── master/                   # Inventarios estáticos (en repo)
│   │   ├── Inventario_MASTER.xlsx
│   │   ├── Inventario_Leaseir_Live.xlsx
│   │   ├── Inventario 30 abril.xlsx
│   │   └── escandallos *.xlsx
│   └── ejercicios/               # Diarios (descargados por workflow)
│       └── ejercicio *.xlsx
├── scripts/
│   ├── download_from_drive.py    # Descarga ejercicios de Drive
│   └── build.sh                  # Orquesta: combina + corre Python + mueve HTML
├── docs/                         # Output servido por GitHub Pages
│   ├── index.html                # Dashboard (Pages lo sirve como página principal)
│   └── Vision_Stock_Mes.html     # Mismo HTML con nombre original
├── requirements.txt
├── .gitignore
├── README.md
└── SETUP.md                      # Este fichero
```
