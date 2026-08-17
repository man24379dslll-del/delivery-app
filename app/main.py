# -*- coding: utf-8 -*-
import os, json, io
from datetime import datetime
from typing import Optional

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import HTMLResponse, JSONResponse
from fastapi.staticfiles import StaticFiles
from fastapi.middleware.cors import CORSMiddleware

from app.analytics import load_and_parse, run_full_analytics
from app.database import save_upload, get_uploads, get_analytics_by_upload

# Лимит 200 МБ на загрузку файлов
app = FastAPI(
    title="Аналитика доставки ТК",
    version="1.0",
)

# Увеличиваем лимит тела запроса через uvicorn — задаётся в Procfile
# Здесь добавляем middleware для больших файлов
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

BASE_DIR = os.path.dirname(os.path.dirname(__file__))
app.mount("/static", StaticFiles(directory=os.path.join(BASE_DIR, "static")), name="static")


@app.get("/", response_class=HTMLResponse)
async def index():
    with open(os.path.join(BASE_DIR, "templates", "index.html"), encoding="utf-8") as f:
        return f.read()


@app.post("/api/upload")
async def upload_file(file: UploadFile = File(...)):
    if not file.filename.endswith(('.xlsx', '.xls')):
        raise HTTPException(400, "Только Excel-файлы (.xlsx, .xls)")

    file_bytes = await file.read()

    try:
        data = load_and_parse(file_bytes)
    except Exception as e:
        raise HTTPException(422, f"Ошибка парсинга файла: {e}")

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
