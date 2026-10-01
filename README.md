# LectorFacturas

Aplicacion web para gestionar facturas, gastos e impuestos con OCR e IA. Incluye dashboard fiscal y exportacion de P&L.

## Requisitos

- Python 3.11+
- Cuenta y clave de OpenAI
- Bucket S3 compatible (AWS, Tigris, Backblaze, etc.)
- PostgreSQL en la nube (Railway, Render, etc.)

## Instalacion local

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Variables de entorno

Configura un archivo `.env` siguiendo el ejemplo en `.env.example` o exporta manualmente:

```bash
export OPENAI_API_KEY="tu_clave"
export DATABASE_URL="postgresql://user:pass@host:5432/dbname"
export STORAGE_BUCKET="mi-bucket"
export STORAGE_REGION="eu-west-1"
export STORAGE_ENDPOINT_URL="https://s3.eu-west-1.amazonaws.com"
export STORAGE_ACCESS_KEY_ID="..."
export STORAGE_SECRET_ACCESS_KEY="..."
export STORAGE_PUBLIC_BASE_URL="https://mi-bucket.s3.eu-west-1.amazonaws.com"
# Bucket privado independiente, solo necesario para la cola persistente.
export PRIVATE_STORAGE_BUCKET="ledged-invoice-analysis-private"
```

Opcionales:

```bash
export OPENAI_CHAT_MODEL="gpt-4o-mini"
# Modelo multimodal para leer la composición visual de facturas y PDFs.
export OPENAI_INVOICE_MODEL="gpt-5.6-sol"
export OPENAI_INVOICE_MAX_OUTPUT_TOKENS="32768"
export OPENAI_INVOICE_TIMEOUT_SECONDS="240"
export OPENAI_INVOICE_REASONING_EFFORT="low"
export OPENAI_INVOICE_AUDIT_REASONING_EFFORT="high"
export OPENAI_MAX_OUTPUT_TOKENS="500"
export ANALYSIS_TIMEOUT_SECONDS="600"
export GUNICORN_TIMEOUT_SECONDS="660"
export MAX_UPLOAD_SIZE_MB="12"
export ACCOUNTING_IMPORT_MAX_ROWS="5000"
# Activalo solo despues de crear el worker descrito mas abajo.
export ASYNC_INVOICE_ANALYSIS_ENABLED="false"
export ASYNC_INVOICE_ANALYSIS_LEASE_SECONDS="720"
export ASYNC_INVOICE_ANALYSIS_RESULT_TTL_SECONDS="86400"
export ANALYSIS_MAX_CONCURRENCY="2"
# Keep these at one for the first production rollout, then evaluate metrics.
export WORKER_CONCURRENCY="1"
export FULL_DOCUMENT_CONCURRENCY="1"
export OCR_CONCURRENCY="1"
export COMPANY_CONCURRENCY="2"
export ASYNC_INVOICE_ANALYSIS_LEASE_RENEWAL_SECONDS="60"
# V2 is a text-only benchmark. Keep it disabled until a controlled test.
export INVOICE_V2_SHADOW_ENABLED="false"
export INVOICE_V2_SHADOW_SAMPLE_RATE="0.10"
export V2_SHADOW_CONCURRENCY="1"
export INVOICE_V2_FAST_TEXT_MAX_CHARS="30000"
```

## Inicializar base de datos

```bash
python init_db.py
```

> La aplicacion tambien crea la base de datos automaticamente al arrancar.

## Arrancar la app

```bash
python app.py
```

Abre `http://127.0.0.1:5000` en el navegador.

## Despliegue en Railway / Render

1. Conecta el repo a Railway o Render.
2. Define las variables de entorno del bloque anterior.
3. Usa el `Procfile` o el `Dockerfile` para el comando de arranque.

El proceso de OCR + IA esta limitado por timeout para evitar bloqueos.

### Cola persistente de facturas

La cola persistente permite seleccionar varias facturas: el navegador entrega los
documentos y un worker los analiza en segundo plano, incluso si se cierra la
pestaña. Esta desactivada por defecto para que el despliegue actual conserve el
comportamiento existente hasta que se provisionen sus recursos.

Para activarla en Render, crea un **Background Worker** con el comando
`python worker.py`. Debe usar la misma `DATABASE_URL`, credenciales del bucket y
variables `OPENAI_*` que el servicio web. Configura en ambos servicios
`ASYNC_INVOICE_ANALYSIS_ENABLED=true`. El bucket debe tener el acceso publico
bloqueado y configurarse como `PRIVATE_STORAGE_BUCKET`; no se reutiliza el bucket
de adjuntos. Los originales se almacenan con una clave aleatoria, no se publica
ninguna URL y se borran al terminar el analisis; el resultado temporal se elimina
como maximo al cabo de 24 horas.

La cola usa una reclamación PostgreSQL con `FOR UPDATE SKIP LOCKED`, un token de
lease por intento y renovaciones periódicas. El resultado solo se acepta si el
worker conserva el mismo token, por lo que un proceso retrasado no puede
sobrescribir el trabajo recuperado por otro. Empieza con `WORKER_CONCURRENCY=1`,
`FULL_DOCUMENT_CONCURRENCY=1` y `OCR_CONCURRENCY=1`. Tras revisar las métricas de
cola, prueba `WORKER_CONCURRENCY=2` manteniendo inicialmente los otros límites en
uno. `COMPANY_CONCURRENCY=2` limita la ocupación de cada empresa para preservar la
equidad entre gestorías.

### Benchmark V2 Fast Path

V2 usa únicamente texto nativo compacto de PDFs digitales y se ejecuta en
**shadow mode**: V1 sigue siendo siempre el único resultado visible y contable.
Por defecto `INVOICE_V2_SHADOW_ENABLED=false`. Cuando se active de forma
controlada, `INVOICE_V2_SHADOW_SAMPLE_RATE` selecciona trabajos de forma estable
por identificador y `V2_SHADOW_CONCURRENCY=1` reserva una única ejecución de baja
prioridad. V2 solo se inicia cuando no hay trabajo V1 activo; no usa OCR ni envía
el PDF, imágenes o texto completo a telemetría. Los resultados comparativos se
guardan en `invoice_analysis_shadow_runs` y el PDF privado se borra tras V2 o al
alcanzar el TTL ya existente.
La variante actual queda identificada como `v2-sol-text-v14`; las ejecuciones
históricas conservan su versión original y la tabla permite comparar varias
variantes del mismo trabajo mediante `job_id + shadow_version`. V14 mantiene
separados la extracción V2, la verificación documental determinista y la
comparación V1/V2 de benchmark. La decisión diagnóstica `accept_v2` o
`fallback_v1` nunca consulta V1: exige PDF digital completo y sin truncar,
documento inequívocamente clasificado como factura, y confirmación documental
de todos los campos críticos aplicables. Incluye identidad fiscal del emisor,
número y fecha de factura, importes, desglose de IVA, retención, otros impuestos,
total, moneda y ecuación contable. Las fechas de pago se verifican por separado:
son informativas y no bloquean una decisión contablemente segura.

El número de factura se prueba primero contra el valor completo propuesto por el
modelo y un contexto local de factura; no acepta prefijos, sufijos ni pedidos o
albaranes. La identidad fiscal del proveedor se verifica contra el documento y
no puede coincidir con la empresa receptora registrada. Los importes requieren
una coincidencia monetaria completa en contexto semántico, no solo una ecuación
que cuadre. Los diagnósticos persistidos son acotados (`status`, método, recuento,
contexto y motivo): nunca guardan texto nativo, fragmentos, prompts, imágenes o
PDFs. V14 sigue siendo exclusivamente shadow: V1 es siempre el único resultado
visible y contable.

V14 conserva las salvaguardas V13 y añade una comprobación espacial efímera basada
en `PyMuPDF.get_text("words")`: agrupa palabras en filas visuales y prueba
relaciones locales etiqueta-valor, columnas de BASE/IVA/TOTAL y filas de IVA. La
geometría sólo puede confirmar una propuesta V2 ya presente en el texto nativo; no
crea campos, no sobrescribe contradicciones, no relaja números de factura y nunca
se persiste. Retención y otros impuestos sólo son `not_applicable` cuando su
ausencia está respaldada por la ecuación contable y no hay un importe etiquetado en
el documento. Tablas ambiguas, NIF no atribuibles, números de factura conflictivos
y contradicciones monetarias siguen forzando `fallback_v1`. Para fechas de dos
dígitos con día y mes ambos menores de 13, conserva el fallback salvo que el propio
documento confirme al proveedor/emisor como español mediante NIF/CIF/NIE o VAT
`ES`. La identidad o dirección española del receptor no cambia el formato de fecha.
En ese caso aplica de forma determinista `DD/MM/YY`; una fecha de vencimiento,
pedido o entrega nunca se reutiliza como fecha de factura.

#### Informe de benchmark persistido

Tras procesar un corpus controlado en shadow mode, ejecuta este informe desde
el Shell del Web Service o Worker de Render. Usa la `DATABASE_URL` ya disponible
en ese entorno y es estrictamente de solo lectura: no llama a OpenAI, no accede
a S3 y no modifica trabajos, métricas ni resultados.

```bash
# Un lote concreto
python scripts/invoice_v2_benchmark.py \
  --version v2-sol-text-v14 \
  --batch-id TU_BATCH_ID

# Las últimas 50 ejecuciones de la variante
python scripts/invoice_v2_benchmark.py --version v2-sol-text-v14 --latest 50

# Un rango de trabajos
python scripts/invoice_v2_benchmark.py \
  --version v2-sol-text-v14 \
  --job-min 40 --job-max 120
```

El informe muestra elegibilidad, validación, `strict_accounting_match`,
`full_document_match`, seguridad contable y calidad de metadata por separado;
además de las decisiones V14, truncamientos, motivos de fallback y campos que
no pudieron verificarse. No imprime PDFs, nombres de archivo, texto extraído,
prompts ni payloads de resultados. `safe_fast_path_candidate`,
`fully_confirmed_fast_path_candidate` y `accept_v2` son métricas de benchmark:
no cambian la ruta oficial V1.

### Descartar un análisis pendiente

`DELETE /api/invoice-analysis-jobs/<job_id>` es un **soft-dismiss**: registra
`dismissed_at` y retira el análisis de la interfaz al recargar. No elimina el
job, su estado, las métricas, los resultados V2 ni el documento temporal antes
de que el worker complete su ciclo normal de limpieza.

## Docker (produccion)

```bash
docker build -t lector-facturas .
docker run -p 8000:8000 --env-file .env lector-facturas
```

## Uso

1. Arrastra facturas o selecciona archivos/carpeta.
2. Completa fecha, proveedor, base imponible y tipo de IVA.
3. Pulsa **Guardar facturas** para registrar todo en PostgreSQL.
4. En **Integraciones**, importa compras o ventas desde CSV/XLSX. El archivo se procesa de forma transitoria, se valida antes de registrar y no se conserva.
5. Filtra por periodo para ver resumenes y graficos.
