# Agente Reporte de Debitos - EVO

Genera un PDF gerencial a partir de los archivos Excel de debitos (hoja `data` o equivalente). Funciona sin Claude: es Python puro.

## Instalacion rapida

### macOS / Linux
```
chmod +x run.sh
./run.sh
```

### Windows
Doble clic en `run.bat`.

El script crea un entorno virtual `.venv`, instala dependencias y ejecuta el agente.

## Uso

Por defecto consulta la API EVO. Con un solo comando obtienes el PDF de una pagina.

```
./run.sh                              # API, todas las sedes
./run.sh --sede "EXITO POBLADO"       # API, una sola sede
./run.sh --list-sedes                 # imprime las sedes disponibles
./run.sh --file ./input/archivo.xlsx  # fallback al archivo Excel
```

### Variables de entorno (opcionales)

```
EVO_DEBITS_URL  = https://action-branches-api.vercel.app/api/debitos
EVO_DEBITS_API_KEY = <api key>
```

Tambien puedes pasar `--api-url` y `--api-key` por linea de comandos.

### Salidas

- `output/Reporte_Debitos[_SEDE]_AAAAMM-AAAAMM.pdf` - reporte de una pagina.
- `output/Motivos_Acciones.md` - catalogo de motivos en espanol con explicacion y accion sugerida.

## Estructura esperada del Excel

El agente busca la hoja `data` (si no, usa la primera). Detecta estas columnas (acepta sinonimos):

| Concepto | Nombres aceptados |
|---|---|
| Fecha del intento | Intento, Fecha, Date |
| Monto | Valor, Monto, Amount |
| Estado | Status, Estado |
| Motivo de rechazo | Motivo del rechazo, Motivo, Reason |
| Tipo de rechazo | Tipo de rechazo, Tipo |
| Sede | Sede/club, Sede, Club |
| Franquicia | Marca de la tarjeta, Franquicia, Marca |

Valores de estado: detecta automaticamente `Aprobado` / `Approved` y `Negado` / `Denied` / `Rechazado`.

## Que incluye el PDF

- Portada con KPIs principales.
- Resumen ejecutivo: tasa de exito, monto recuperado, monto en riesgo.
- Conclusiones clave en lenguaje natural.
- Grafico donut (aprobado vs negado).
- Pie chart de motivos de rechazo (Top + Otros).
- Tendencia diaria con tasa de exito.
- Bar chart de las 15 sedes con mas fallos.
- Bar chart de las 15 sedes con peor tasa de exito.
- Tabla detallada por sede (resaltando sedes criticas).
- Top 20 motivos de rechazo con frecuencia y porcentaje.
- Desempeno por franquicia.
- Recomendaciones accionables.

## Requisitos

- Python 3.9 o superior.
- Conexion a internet la primera vez (para `pip install`).

Dependencias (ver `requirements.txt`): pandas, numpy, openpyxl, matplotlib, reportlab.

## Programacion automatica (opcional)

### macOS / Linux (cron)
Editar con `crontab -e`:
```
0 9 1 * * /ruta/a/agent_report_debits/run.sh >> /ruta/a/agent_report_debits/output/last_run.log 2>&1
```
Corre el primer dia de cada mes a las 9:00.

### Windows (Programador de Tareas)
Crea una tarea que ejecute `run.bat` con la frecuencia deseada.

## Solucion de problemas

- `ERROR: no se encontraron archivos .xlsx` -> Coloca el archivo en `input/` o pasa `--input`.
- Error en `pip install` -> Verifica que Python esta en el PATH; en Windows reinstala marcando "Add to PATH".
- PDF vacio o sin graficos -> Asegurate que la hoja se llama `data` y tiene las columnas esperadas.
