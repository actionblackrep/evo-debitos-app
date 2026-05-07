# Deploy - paso a paso (apps de equipo y de gerentes)

Hay dos apps en este repo, comparten codigo:

| App | Archivo | Quien la usa |
|---|---|---|
| Estandar (admin) | `app.py` | Equipo operativo |
| Gerentes Generales | `app_managers.py` | Gerentes Generales |

Ambas se deployan a Streamlit Community Cloud gratis. URLs separadas.

La app de **Gerentes Generales** intenta usar la API EVO directamente. Si la
API no responde, lee automaticamente el ultimo dataset publicado por la app
de **admin** (un parquet en una rama del repo de GitHub). Asi los gerentes
siempre ven datos sin tener que subir nada.

## Paso 1. Cuenta de GitHub
Si ya tienes una, salta al paso 2.
1. https://github.com/signup
2. Crea cuenta y verifica correo.

## Paso 2. Crear repo
1. Boton "+" -> "New repository".
2. Nombre: `evo-debitos-app`.
3. Visibilidad: Private.
4. NO marques "Add a README".
5. "Create repository".

## Paso 3. Subir codigo

Sube TODOS estos archivos del directorio `agent_report_debits`:
- `app.py`
- `app_managers.py`
- `data_loader.py`
- `shared_cache.py`
- `generate_report.py`
- `motivos.py`
- `requirements.txt`
- `runtime.txt`
- `.streamlit/config.toml`

### Opcion A (sin terminal)
"uploading an existing file" -> arrastra los archivos -> commit.

### Opcion B (con terminal)
```
cd "ruta/a/agent_report_debits"
git init
git add app.py app_managers.py data_loader.py shared_cache.py \
        generate_report.py motivos.py requirements.txt runtime.txt .streamlit
git commit -m "primer deploy"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/evo-debitos-app.git
git push -u origin main
```

## Paso 4. Crear PAT (token) para el cache compartido

El admin publica el ultimo dataset en una rama del repo de GitHub via API.
Necesita un Personal Access Token. Tienes dos opciones, **escoge UNA**.

### Opcion A - Fine-grained PAT (mas seguro)

1. URL exacta: https://github.com/settings/personal-access-tokens/new
   (asegurate que diga `personal-access-tokens`, NO `tokens/new`).
2. Token name: `evo-debitos-data-sync`. Expiration: 1 year.
3. Resource owner: tu cuenta.
4. Repository access: **"Only select repositories"** -> elige `evo-debitos-app`.
   Si dejas "Public Repositories (read-only)" la seccion de permisos NO aparece.
5. Baja hasta "Repository permissions" (lista larga, despues de elegir el repo):
   - **Contents: Read and write**
6. "Generate token" -> copia el token (empieza con `github_pat_`).

### Opcion B - Classic PAT (mas simple, funciona igual)

Usalo si la opcion A no muestra la seccion de permisos o quieres ir mas rapido.

1. URL: https://github.com/settings/tokens/new
2. Note: `evo-debitos-data-sync`. Expiration: 1 year.
3. Scopes: marca **`repo`** (control total sobre repos privados).
   Esto cubre contents y todo lo necesario.
4. "Generate token" -> copia el token (empieza con `ghp_`).

Ambos formatos se usan igual en los secrets:
```
GITHUB_TOKEN = "github_pat_..."   # o "ghp_..." si elegiste Opcion B
```

## Paso 5. Crear la rama de cache

Una sola vez. Tres rutas, **escoge una**.

### Opcion 0 - Saltar este paso (recomendado)
El helper `shared_cache.ensure_branch()` crea la rama automaticamente la
primera vez que el admin cargue datos, siempre que el token del Paso 4 tenga
los permisos correctos. Si quieres ir directo al Paso 6, puedes saltar esto.

### Opcion A - URL directa de Branches
1. Abre `https://github.com/TU_USUARIO/evo-debitos-app/branches`
2. Boton verde "New branch" arriba a la derecha.
3. Branch name: `data-cache`. Source: `main`.
4. "Create branch".

### Opcion B - Dropdown del repo (solo si ya hay archivos en `main`)
1. Pestana "Code" del repo (pagina principal).
2. Arriba del listado de archivos, lado izquierdo, boton con icono de rama
   que dice `main` con flecha hacia abajo.
3. Si el repo esta vacio este boton NO aparece. Sube primero los archivos
   del Paso 3 y vuelve.
4. Click en el boton -> en el cuadro de busqueda escribe `data-cache` ->
   click en "Create branch: data-cache from main".

### Opcion C - Terminal
```
git checkout --orphan data-cache
git rm -rf .
echo "Cache compartido para EVO Debitos" > README.md
git add README.md
git commit -m "init data-cache branch"
git push origin data-cache
git checkout main
```

## Paso 6. Deploy app admin (`app.py`)

1. https://share.streamlit.io -> "Sign in with GitHub".
2. "New app".
3. Repository: `TU_USUARIO/evo-debitos-app`. Branch: `main`. Main file: `app.py`.
4. App URL: `evo-debitos`.
5. Advanced settings -> Secrets:
   ```
   EVO_DEBITS_URL = "https://action-branches-api.vercel.app/api/debitos"
   EVO_DEBITS_API_KEY = "dLjaU5u4LfycyRpbBTU7EMcXDBL2zFrOiX6fBWO6b-s"
   GITHUB_TOKEN = "github_pat_..."
   GITHUB_REPO = "TU_USUARIO/evo-debitos-app"
   GITHUB_DATA_BRANCH = "data-cache"
   GITHUB_DATA_PATH = "data/latest.parquet"
   ```
6. "Deploy".

## Paso 7. Deploy app de Gerentes Generales (`app_managers.py`)

Repite el paso 6 con:
- Main file path: `app_managers.py`
- App URL: `evo-debitos-gm`
- Secrets: los MISMOS del paso 6 (incluyendo `GITHUB_TOKEN` para leer el cache).

Resultado:
- https://evo-debitos.streamlit.app/        (admin)
- https://evo-debitos-gm.streamlit.app/     (gerentes)

## Paso 8. Restringir acceso (recomendado)
En cada app: Settings -> Sharing -> "Only specific people" -> agregar correos.

## Como funciona en la practica

1. Admin entra a `evo-debitos`, carga desde API o sube Excel.
2. Tras la carga, la app publica un parquet en `data-cache/data/latest.parquet`
   automaticamente. Aparece la nota "🛰  Publicado para Gerentes Generales".
3. Cada gerente abre `evo-debitos-gm`. La app intenta la API:
   - Si responde, usa la API.
   - Si no, lee el parquet publicado por el admin.
4. El gerente elige su sede. Solo los datos de esa sede se procesan.

## Local development

```
cd agent_report_debits
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app_managers.py      # o app.py
```

Variables de entorno (si quieres probar el cache compartido localmente):
```
export EVO_DEBITS_URL="https://action-branches-api.vercel.app/api/debitos"
export EVO_DEBITS_API_KEY="..."
export GITHUB_TOKEN="github_pat_..."
export GITHUB_REPO="TU_USUARIO/evo-debitos-app"
```

## Actualizar codigo
Cualquier push a `main` redeploya automaticamente las dos apps.
La rama `data-cache` NO triggerea redeploy (el deploy esta en `main`).

## Troubleshooting

| Sintoma | Causa | Fix |
|---|---|---|
| GM dice "API no respondio y admin no ha publicado" | Falta primer push del admin | Que el admin entre a `evo-debitos` y cargue 1 vez |
| Admin no muestra "🛰  Publicado..." | `GITHUB_TOKEN` o `GITHUB_REPO` mal configurados | Revisa secrets en `share.streamlit.io` -> Settings |
| GM siempre usa datos viejos | Cache TTL local | Boton "Forzar reconexion API" en sidebar |
| 403/404 al publicar | PAT sin permisos `Contents:write` | Regenera PAT con permisos correctos |
| ModuleNotFoundError: pyarrow | requirements desactualizado | Sube `requirements.txt` actualizado |

## Costo
Streamlit Community Cloud: gratis. GitHub: gratis (repo private). Sin tarjetas.
