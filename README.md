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
```

Opcionales:

```bash
export OPENAI_CHAT_MODEL="gpt-4o-mini"
export OPENAI_MAX_OUTPUT_TOKENS="500"
export ANALYSIS_TIMEOUT_SECONDS="120"
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

## Docker (produccion)

```bash
docker build -t lector-facturas .
docker run -p 8000:8000 --env-file .env lector-facturas
```

## Uso

1. Arrastra facturas o selecciona archivos/carpeta.
2. Completa fecha, proveedor, base imponible y tipo de IVA.
3. Pulsa **Guardar facturas** para registrar todo en PostgreSQL.
4. Filtra por periodo para ver resumenes y graficos.
