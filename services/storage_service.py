import logging
import os
from typing import Optional

import boto3
from botocore.config import Config

logger = logging.getLogger(__name__)

_client = None


def _get_client():
    global _client
    if _client is not None:
        return _client

    region = os.getenv("STORAGE_REGION", "us-east-1")
    endpoint_url = os.getenv("STORAGE_ENDPOINT_URL")
    access_key = os.getenv("STORAGE_ACCESS_KEY_ID")
    secret_key = os.getenv("STORAGE_SECRET_ACCESS_KEY")

    _client = boto3.client(
        "s3",
        region_name=region,
        endpoint_url=endpoint_url,
        aws_access_key_id=access_key,
        aws_secret_access_key=secret_key,
        config=Config(signature_version="s3v4"),
    )
    return _client


def _build_public_url(bucket: str, key: str) -> str:
    public_base = os.getenv("STORAGE_PUBLIC_BASE_URL")
    if public_base:
        return f"{public_base.rstrip('/')}/{key}"

    endpoint_url = os.getenv("STORAGE_ENDPOINT_URL")
    region = os.getenv("STORAGE_REGION", "us-east-1")
    if endpoint_url:
        return f"{endpoint_url.rstrip('/')}/{bucket}/{key}"

    return f"https://{bucket}.s3.{region}.amazonaws.com/{key}"


def get_public_url(key: str) -> str:
    bucket = os.getenv("STORAGE_BUCKET")
    if not bucket:
        raise RuntimeError("STORAGE_BUCKET no configurado")
    return _build_public_url(bucket, key)


def upload_bytes(data: bytes, key: str, content_type: Optional[str] = None) -> str:
    bucket = os.getenv("STORAGE_BUCKET")
    if not bucket:
        raise RuntimeError("STORAGE_BUCKET no configurado")

    client = _get_client()
    extra_args = {}
    if content_type:
        extra_args["ContentType"] = content_type

    client.put_object(Bucket=bucket, Key=key, Body=data, **extra_args)
    url = _build_public_url(bucket, key)
    logger.info("Archivo subido a storage: %s", url)
    return url
