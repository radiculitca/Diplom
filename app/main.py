from fastapi import FastAPI, Request, UploadFile, File
from typing import List
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse, Response
import pandas as pd
import os

import json
import asyncio
from fastapi.responses import StreamingResponse

from app.data_logic import clean_dataframe, generate_report_data, get_column_groups, is_system_column
from app.schemas import ProcessSheetsRequest, AnalyzeRequest, ExportDocxRequest
from app.docx_gen import generate_docx

app = FastAPI(title="Система аналитики опросов МГУ им. Огарева")

app.mount("/static", StaticFiles(directory="app/static"), name="static")
templates = Jinja2Templates(directory="app/templates")

UPLOAD_DIR = "uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

@app.get("/")
async def read_root(request: Request):
    return templates.TemplateResponse(request=request, name="index.html", context={"title": "Аналитика опросов"})

@app.post("/upload")
async def upload_files(files: List[UploadFile] = File(...)):
    try:
        response_data = []
        for file in files[:10]:
            contents = await file.read()
            file_ext = file.filename.split('.')[-1].lower()
            safe_filename = f"raw_{os.path.basename(file.filename)}".replace(" ", "_")
            filepath = os.path.join(UPLOAD_DIR, safe_filename)

            with open(filepath, "wb") as f:
                f.write(contents)

            if file_ext in ['csv', 'tsv']:
                sheet_names = ["CSV Данные"]
            else:
                xl = pd.ExcelFile(filepath)
                sheet_names = xl.sheet_names

            response_data.append({
                "original_name": file.filename,
                "filename": safe_filename,
                "sheets": sheet_names
            })

        return {"files": response_data}
    except Exception as e:
        return JSONResponse(status_code=400, content={"message": f"Ошибка загрузки файлов: {str(e)}"})

@app.post("/process_sheets")
async def process_sheets(request: ProcessSheetsRequest):
    try:
        response_data = []

        for file_data in request.files:
            filepath = os.path.join(UPLOAD_DIR, file_data.filename)
            if not os.path.exists(filepath):
                continue

            if file_data.filename.endswith('.csv') or file_data.filename.endswith('.tsv'):
                try: df = pd.read_csv(filepath, sep='\t')
                except Exception: df = pd.read_csv(filepath)
            else:
                dfs = []
                for sheet in file_data.sheets:
                    dfs.append(pd.read_excel(filepath, sheet_name=sheet))
                df = pd.concat(dfs, ignore_index=True)

            df_clean = clean_dataframe(df)

            clean_filename = f"clean_{file_data.filename}.parquet"
            clean_filepath = os.path.join(UPLOAD_DIR, clean_filename)
            df_clean.to_parquet(clean_filepath, index=False)

            groups = get_column_groups(df_clean.columns)
            columns_data = [{"name": q, "is_system": is_system_column(q)} for q in groups.keys()]

            response_data.append({
                "original_name": file_data.filename.replace("raw_", ""),
                "clean_filename": clean_filename,
                "columns": columns_data
            })

        return {"processed_files": response_data}
    except Exception as e:
        return JSONResponse(status_code=400, content={"message": f"Ошибка обработки листов: {str(e)}"})

@app.post("/analyze")
async def analyze_data(request: AnalyzeRequest):
    results = generate_report_data(UPLOAD_DIR, request)
    return {"results": results}

@app.post("/export_docx_stream")
async def export_docx_stream(request: ExportDocxRequest):
    import base64

    questions = [q.model_dump() for q in request.questions]

    async def event_generator():
        loop = asyncio.get_event_loop()
        queue = asyncio.Queue()

        def progress_cb(current, total, label):
            loop.call_soon_threadsafe(
                queue.put_nowait,
                {"type": "progress", "current": current, "total": total, "label": label}
            )

        def run_generate():
            try:
                data_bytes, analysis_bytes = generate_docx(questions, progress_callback=progress_cb)
                loop.call_soon_threadsafe(queue.put_nowait, {
                    "type": "done",
                    "data": base64.b64encode(data_bytes).decode(),
                    "analysis": base64.b64encode(analysis_bytes).decode()
                })
            except Exception as e:
                loop.call_soon_threadsafe(queue.put_nowait, {"type": "error", "message": str(e)})

        import threading
        threading.Thread(target=run_generate, daemon=True).start()

        while True:
            msg = await queue.get()
            yield f"data: {json.dumps(msg, ensure_ascii=False)}\n\n"
            if msg["type"] in ("done", "error"):
                break

    return StreamingResponse(event_generator(), media_type="text/event-stream")


@app.post("/export_docx")
async def export_docx(request: ExportDocxRequest):
    """Синхронная версия (запасная) — возвращает только файл данных."""
    try:
        data_bytes, _ = generate_docx([q.model_dump() for q in request.questions])
        return Response(
            content=data_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=report_data.docx"}
        )
    except Exception as e:
        return JSONResponse(status_code=400, content={"message": f"Ошибка генерации документа: {str(e)}"})
