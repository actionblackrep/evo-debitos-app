# Deploy - paso a paso (apps de equipo y de gerentes)

Hay dos apps en este repo, comparten codigo:

| App | Archivo | Quien la usa |
|---|---|---|
| Estandar | `app.py` | Equipo operativo (filtra despues de cargar) |
| Gerentes | `app_managers.py` | Gerentes Generales (sede obligatoria, mas rapida) |

Ambas se deployan a Streamlit Community Cloud gratis. URLs separadas.

## Paso 1. Cuenta de GitHub
Si ya tienes una, salta al paso 2.
1. https://github.com/signup
2. Crea cuenta y verifica correo.

## Paso 2. Crear repo
1. Boton "+" arriba a la derecha -> "New repository".
2. Nombre: `evo-debitos-app`.
3. Visibilidad: Private.
4. NO marques "Add a README".
5. "Create repository".

## Paso 3. Subir archivos

Sube TODOS estos archivos del directorio `agent_report_debits`:
- `app.py`
- `app_managers.py`
- `data_loader.py`
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
git add app.py app_managers.py data_loader.py generate_report.py motivos.py \
        requirements.txt runtime.txt .streamlit
git commit -m "primer deploy"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/evo-debitos-app.git
git push -u origin main
```

## Paso 4. Deploy app estandar
1. https://share.streamlit.io -> "Sign in with GitHub".
2. "New app".
3. Repository: `TU_USUARIO/evo-debitos-app`. Branch: `main`. Main file: `app.py`.
4. App URL: `evo-debitos`.
5. Advanced settings -> Secrets:
   ```
   EVO_DEBITS_URL = "https://action-branches-api.vercel.app/api/debitos"
   EVO_DEBITS_API_KEY = "dLjaU5u4LfycyRpbBTU7EMcXDBL2zFrOiX6fBWO6b-s"
   ```
6. "Deploy".

## Paso 5. Deploy app de Gerentes Generales
Repite el paso 4 cambiando:
- Main file path: `app_managers.py`
- App URL: `evo-debitos-gm`

Resultado: dos URLs publicas. Ej:
- https://evo-debitos.streamlit.app/
- https://evo-debitos-gm.streamlit.app/

## Paso 6. Restringir acceso (recomendado)
En cada app: Settings -> Sharing -> "Only specific people" -> agregar correos.

## Paso 7. Compartir
- Equipo operativo: URL de `evo-debitos`.
- Gerentes Generales: URL de `evo-debitos-gm`.

## Local development

```
cd agent_report_debits
python3 -m venv .venv
source .venv/bin/activate          # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app_managers.py      # o app.py
```

Variables de entorno opcionales:
```
export EVO_DEBITS_URL="https://action-branches-api.vercel.app/api/debitos"
export EVO_DEBITS_API_KEY="..."
```

## Actualizar codigo
Cualquier push a `main` redeploya automaticamente las dos apps.

## Troubleshooting

| Sintoma | Causa | Fix |
|---|---|---|
| "ModuleNotFoundError: pyarrow" | requirements desactualizado | Sube `requirements.txt` actualizado |
| App tarda > 30s en GM | primer indexado del Excel | Normal, queda cacheado para todas las sedes |
| "No fue posible cargar los datos" | API caida | Cambia a "Subir Excel" en la sidebar |
| Memoria al limite | archivo > 60 MB | Usa la API o sube en partes |

## Costo
Streamlit Community Cloud: gratis. Sin tarjeta de credito.
