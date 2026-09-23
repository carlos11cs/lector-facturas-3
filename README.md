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
