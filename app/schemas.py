from typing import List, Dict, Any
from pydantic import BaseModel

class ProcessSheetFile(BaseModel):
    filename: str
    sheets: List[str]

class ProcessSheetsRequest(BaseModel):
    files: List[ProcessSheetFile]

class ColumnConfig(BaseModel):
    column: str
    viz_type: List[str]
    file_mapping: Dict[str, str]

class AnalyzeRequest(BaseModel):
    file_labels: Dict[str, str]
    file_colors: Dict[str, str]
    configs: List[ColumnConfig]

class ExportRow(BaseModel):
    answer: str
    counts: Dict[str, int]

class ExportQuestion(BaseModel):
    table_num: int
    question_name: str
    h1: str
    h2: str
    h3: str
    file_keys: List[str]
    file_labels: Dict[str, str]
    rows: List[ExportRow]
    file_totals: Dict[str, int]
    show_total: bool

class ExportDocxRequest(BaseModel):
    questions: List[ExportQuestion]
