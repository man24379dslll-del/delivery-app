# -*- coding: utf-8 -*-
import os, json, io
import pandas as pd
from datetime import datetime

from fastapi import FastAPI, UploadFile, File, HTTPException, Query
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from app.analytics import load_and_parse, run_full_analytics
from app.database import save_upload, get_uploads, get_analytics_by_upload

app = FastAPI(title="Аналитика доставки ТК", version="1.0")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")

# Кэш последнего загруженного файла (в памяти процесса)
_cached_data = None
_cached_filename = None


@app.get("/", response_class=HTMLResponse)
async def index():
    with open(os.path.join(BASE_DIR, "templates", "index.html"), encoding="utf-8") as f:
        content = f.read()
    from fastapi.responses import Response
    return Response(
        content=content,
        media_type="text/html",
        headers={
            "Content-Security-Policy": (
                "default-src 'self'; "
                "script-src 'self' 'unsafe-inline' 'unsafe-eval' https://cdn.jsdelivr.net https://fonts.googleapis.com; "
                "style-src 'self' 'unsafe-inline' https://fonts.googleapis.com; "
                "font-src 'self' https://fonts.gstatic.com; "
                "img-src 'self' data:; "
                "connect-src 'self';"
            )
        }
    )


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    global _cached_data, _cached_filename

    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "Только Excel-файлы (.xlsx, .xls)")

    file_bytes = await file.read()

    try:
        data = load_and_parse(file_bytes)
    except Exception as e:
        raise HTTPException(422, f"Ошибка парсинга файла: {e}")

    # Кэшируем для перефильтрации по месяцу
    _cached_data = data
    _cached_filename = file.filename

    analytics = run_full_analytics(data)

    upload_id = None
    try:
        upload_id = await save_upload(
            filename=file.filename,
            rows_count=len(data),
            analytics=analytics,
        )
    except Exception as e:
        print(f"[DB skip] {e}")

    return JSONResponse({
        "status":    "ok",
        "upload_id": upload_id,
        "filename":  file.filename,
        "rows":      len(data),
        "analytics": analytics,
    })


@app.get("/api/filter")
async def filter_by_month(month: str = Query(default="")):
    """Перефильтровывает аналитику по месяцу без повторной загрузки файла."""
    global _cached_data, _cached_filename
    if _cached_data is None:
        raise HTTPException(400, "Сначала загрузите файл")
    try:
        analytics = run_full_analytics(_cached_data, month=month)
        return JSONResponse({
            "status":   "ok",
            "filename": _cached_filename,
            "rows":     len(_cached_data),
            "analytics": analytics,
        })
    except Exception as e:
        raise HTTPException(500, str(e))


@app.get("/api/uploads")
async def list_uploads():
    try:
        uploads = await get_uploads()
        return JSONResponse({"uploads": uploads})
    except Exception as e:
        return JSONResponse({"uploads": [], "error": str(e)})


@app.get("/api/uploads/{upload_id}")
async def get_upload(upload_id: str):
    try:
        result = await get_analytics_by_upload(upload_id)
        if not result:
            raise HTTPException(404, "Загрузка не найдена")
        return JSONResponse(result)
    except HTTPException:
        raise
    except Exception as e:
        raise HTTPException(500, str(e))


@app.get("/api/health")
async def health():
    return {"status": "ok", "time": datetime.now().isoformat()}


# ── Google Sheets интеграция ────────────────────────────────────────
from app.google_sheets import is_configured, get_cached_or_load, SPREADSHEET_ID

@app.get("/api/gsheets/status")
async def gsheets_status():
    """Проверяет статус подключения к Google Sheets."""
    configured = is_configured()
    return JSONResponse({
        "configured": configured,
        "spreadsheet_id": SPREADSHEET_ID,
        "cached": _cached_data is not None,
        "message": "Google Sheets подключён" if configured else "Нужен credentials.json"
    })


@app.post("/api/gsheets/sync")
async def gsheets_sync(force: bool = False):
    """Загружает данные из Google Sheets и пересчитывает аналитику."""
    global _cached_data, _cached_filename
    if not is_configured():
        raise HTTPException(400, "Google Sheets не настроен. Добавьте credentials.json")
    try:
        from app.analytics import load_and_parse, run_full_analytics
        import io

        print("[GSheets Sync] Загружаем данные...")
        raw_df = get_cached_or_load(force_reload=force)

        # Конвертируем DataFrame обратно в Excel bytes для load_and_parse
        # (переиспользуем существующий парсер)
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='openpyxl') as writer:
            # Группируем по источнику
            for src, grp in raw_df.groupby('_источник', dropna=False):
                sheet = str(src) if pd.notna(src) else 'Sheet1'
                grp.drop(columns=['_источник'], errors='ignore').to_excel(writer, sheet_name=sheet[:31], index=False)
        buf.seek(0)

        data = load_and_parse(buf.read())
        analytics = run_full_analytics(data)

        _cached_data = data
        _cached_filename = f"Google Sheets · {SPREADSHEET_ID[:8]}..."

        return JSONResponse({
            "status": "ok",
            "rows": len(data),
            "filename": _cached_filename,
            "analytics": analytics,
        })
    except Exception as e:
        raise HTTPException(500, f"Ошибка синхронизации: {e}")
