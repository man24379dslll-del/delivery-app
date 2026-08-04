# -*- coding: utf-8 -*-
"""
Работа с Supabase (PostgreSQL).
Если переменные окружения не заданы — модуль работает в режиме заглушки
(данные не сохраняются, но приложение не падает).
"""
import os, json
from datetime import datetime
from typing import Optional

SUPABASE_URL = os.getenv("SUPABASE_URL", "")
SUPABASE_KEY = os.getenv("SUPABASE_KEY", "")

_client = None

def _get_client():
    global _client
    if _client is not None:
        return _client
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise RuntimeError("Supabase не настроен. Задайте SUPABASE_URL и SUPABASE_KEY в .env")
    from supabase import create_client
    _client = create_client(SUPABASE_URL, SUPABASE_KEY)
    return _client


# ── SQL для инициализации таблиц (запустить один раз в Supabase) ──────────
INIT_SQL = """
CREATE TABLE IF NOT EXISTS uploads (
    id          UUID PRIMARY KEY DEFAULT gen_random_uuid(),
    filename    TEXT NOT NULL,
    uploaded_at TIMESTAMPTZ DEFAULT NOW(),
    rows_count  INTEGER,
    analytics   JSONB
);
"""


async def save_upload(filename: str, rows_count: int, analytics: dict) -> Optional[str]:
    """Сохраняет результаты загрузки в Supabase. Возвращает id записи."""
    client = _get_client()
    response = client.table("uploads").insert({
        "filename":   filename,
        "rows_count": rows_count,
        "analytics":  analytics,
    }).execute()
    data = response.data
    if data:
        return data[0]["id"]
    return None


async def get_uploads() -> list[dict]:
    """Возвращает список всех загрузок (без аналитики — только мета)."""
    client = _get_client()
    response = (
        client.table("uploads")
        .select("id, filename, uploaded_at, rows_count")
        .order("uploaded_at", desc=True)
        .limit(50)
        .execute()
    )
    return response.data or []


async def get_analytics_by_upload(upload_id: str) -> Optional[dict]:
    """Возвращает сохранённую аналитику по id загрузки."""
    client = _get_client()
    response = (
        client.table("uploads")
        .select("*")
        .eq("id", upload_id)
        .single()
        .execute()
    )
    return response.data
