# Deploy de la app web - paso a paso

Lo que vas a obtener: una URL tipo `https://evo-debitos.streamlit.app/` que cualquier persona en Windows o macOS abre en su navegador. No hay que instalar nada en su computador.

Tiempo estimado: 8-10 minutos. Solo se hace una vez.

## Paso 1. Tener cuenta de GitHub

Si ya tienes una, salta al paso 2.

1. Entra a [https://github.com/signup](https://github.com/signup)
2. Crea tu cuenta gratis (correo + contrasena).
3. Verifica tu correo.

## Paso 2. Crear un repositorio nuevo

1. Una vez dentro de GitHub, clic en el boton verde "+" arriba a la derecha → "New repository".
2. Nombre del repo: `evo-debitos-app` (o el que quieras).
3. Visibilidad: **Private** (recomendado, asi solo quien tu invites puede entrar).
4. NO marques "Add a README".
5. Clic en "Create repository".

## Paso 3. Subir los archivos de la carpeta `agent_report_debits`

GitHub te muestra una pagina con instrucciones. Hay dos caminos:

### Opcion A (mas facil, sin terminal)

1. En la pagina del repo recien creado, clic en "uploading an existing file".
2. Arrastra estos archivos a la ventana del navegador:
   - `app.py`
   - `generate_report.py`
   - `motivos.py`
   - `requirements.txt`
   - `runtime.txt`
   - Carpeta `.streamlit/` (con `config.toml` adentro)
3. Abajo, escribe en el cuadro de mensaje: "primer deploy"
4. Clic en "Commit changes".

### Opcion B (con terminal, si te sientes comodo)

```
cd "ruta/a/agent_report_debits"
git init
git add app.py generate_report.py motivos.py requirements.txt runtime.txt .streamlit
git commit -m "primer deploy"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/evo-debitos-app.git
git push -u origin main
```

## Paso 4. Conectar Streamlit Community Cloud

Streamlit Community Cloud es **gratis** y aloja la app por ti.

1. Entra a [https://share.streamlit.io](https://share.streamlit.io)
2. Clic en "Sign in with GitHub". Autoriza el acceso.
3. Clic en "New app".
4. Llena:
   - **Repository**: `TU_USUARIO/evo-debitos-app`
   - **Branch**: `main`
   - **Main file path**: `app.py`
   - **App URL** (opcional): elige un subdominio, ej. `evo-debitos`
5. Antes de hacer "Deploy", abre "Advanced settings" y agrega los secrets:

```
EVO_DEBITS_URL = "https://action-branches-api.vercel.app/api/debitos"
EVO_DEBITS_API_KEY = "dLjaU5u4LfycyRpbBTU7EMcXDBL2zFrOiX6fBWO6b-s"
```

6. Clic en "Deploy". En 2-3 minutos veras la URL publica.

## Paso 5. Compartir el enlace

Manda la URL (`https://evo-debitos.streamlit.app/` o la que hayas elegido) al equipo.

Para usarla:

1. Abren el enlace en cualquier navegador (Chrome, Safari, Edge, Firefox).
2. Eligen la sede en el dropdown.
3. Eligen el rango de fechas.
4. Clic en "Generar PDF" → descarga.

No hay que instalar Python, no hay que correr terminal, no hay que actualizar nada.

## Restringir acceso (opcional, recomendado)

Streamlit Community Cloud permite restringir la app por correo:

1. Ya desplegada la app, en la pagina del proyecto en `share.streamlit.io`, abre "Settings" → "Sharing".
2. Cambia a "Only specific people".
3. Agrega los correos del equipo.

Solo esos correos podran abrir la URL despues de hacer login con Google.

## Actualizar el codigo

Cuando quieras modificar algo:

1. Sube los cambios al repo de GitHub (otra vez "Add file" → "Upload files", o `git push`).
2. Streamlit detecta el cambio y redespliega automaticamente en 1-2 minutos.

## Alternativas de hosting

- **Render.com**: similar, gratis para apps pequenas. Conectas el repo igual.
- **Hugging Face Spaces**: gratis, aceptan apps Streamlit. URL tipo `huggingface.co/spaces/TU_USUARIO/...`.
- **Self-host**: clona el repo en un servidor con Python y corre `streamlit run app.py`. Recomendado solo si el equipo de TI quiere mantenerlo.

## Si la app no levanta

Errores comunes y solucion:

| Sintoma | Causa | Solucion |
|---|---|---|
| "ModuleNotFoundError: streamlit" | falta `requirements.txt` | Sube `requirements.txt` al repo |
| "No fue posible cargar los datos" | endpoint API caido | En la barra lateral, cambia a "Subir Excel" y sube el archivo |
| "ImportError: motivos" | falta `motivos.py` | Sube `motivos.py` al repo |
| App tarda mucho | primera carga es lenta | Espera 30-60 segundos. Despues cachea |

## Costo

Streamlit Community Cloud es **gratis para apps publicas o con pocos usuarios** (limite generoso para uso interno). Si llegan a necesitar mas recursos, el plan pagado empieza en USD 250/mes pero el equipo de EVO seguramente no lo necesita.
