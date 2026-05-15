from typing import List, Dict, Any, Optional
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
    merged_columns: List[str] = []

class AnalyzeRequest(BaseModel):
    file_labels: Dict[str, str]
    file_colors: Dict[str, str]
    configs: List[ColumnConfig]

class ExportRow(BaseModel):
    answer: str
    counts: Dict[str, int]

class SectionInfo(BaseModel):
    name: str
    description: str = ""
    color: str = ""

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
    section: Optional[SectionInfo] = None
    viz_tab: Optional[str] = None        # 'table' | 'bar' | 'stacked' | 'pie' | None
    chart_direction: str = 'y'           # 'y' = column, 'x' = horizontal bar
    show_legend: bool = True
    hidden_col: str = 'none'             # 'none' | 'count' | 'percent'

class ExportDocxRequest(BaseModel):
    questions: List[ExportQuestion]

class AiGroupRequest(BaseModel):
    answers: List[str]
    question_name: str
