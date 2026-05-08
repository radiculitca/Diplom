from fastapi import FastAPI, Request, UploadFile, File
from typing import List
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import JSONResponse, Response
import pandas as pd
import os

from app.data_logic import clean_dataframe, generate_report_data, get_column_groups, is_system_column
from app.schemas import ProcessSheetsRequest, AnalyzeRequest, ExportDocxRequest
from app.docx_gen_all import generate_docx

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

@app.post("/export_docx")
async def export_docx(request: ExportDocxRequest):
    try:
        docx_bytes = generate_docx([q.model_dump() for q in request.questions])
        return Response(
            content=docx_bytes,
            media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            headers={"Content-Disposition": "attachment; filename=report.docx"}
        )
    except Exception as e:
        return JSONResponse(status_code=400, content={"message": f"Ошибка генерации документа: {str(e)}"})
