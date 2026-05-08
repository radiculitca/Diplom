import pandas as pd
import re
import os
from datetime import datetime


def clean_column_name(col_name):
    name = str(col_name).strip()
    name = re.sub(r"^\s*\d+\s*[\)\.]\s*", "", name)
    name = re.sub(r"\s*:\s*$", "", name)
    return name


def clean_answer_text(answer):
    if pd.isna(answer):
        return "Нет ответа"
    ans = str(answer).strip()
    ans = re.sub(r"^\s*\d+\s*[\)\.]\s*", "", ans)
    return ans


def is_system_column(col_name):
    c = str(col_name).lower().strip()
    sys_exact = [
        "id",
        "айди",
        "timestamp",
        "email",
        "почта",
        "адрес электронной почты",
        "время",
        "дата",
        "time",
        "date",
    ]
    sys_contains = [
        "время создания",
        "дата создания",
        "время заполнения",
        "дата и время",
        "время начала",
        "время завершения",
        "completion time",
        "start time",
        "время изменения",
        "дата изменения",
        "отметка времени",
    ]
    if c in sys_exact:
        return True
    for kw in sys_contains:
        if kw in c:
            return True
    return False


def unify_numbered_answers(series):
    if series.dropna().empty:
        return series
    s_str = series.astype(str).str.strip()
    extracted = s_str.str.extract(r"^(\d+)[\)\.]\s*(.*)")

    if extracted[0].notna().sum() > 0:
        mapping = {}
        for num in extracted[0].dropna().unique():
            texts = extracted[extracted[0] == num][1].str.strip()
            texts = texts[texts != ""]
            if not texts.empty:
                russian_texts = texts[
                    texts.str.contains(r"[А-Яа-яЁё]", regex=True, na=False)
                ]
                if not russian_texts.empty:
                    best_text = russian_texts.mode().iloc[0]
                else:
                    best_text = texts.mode().iloc[0]
                mapping[num] = f"{num}) {best_text}"

        def apply_map(val):
            if pd.isna(val) or str(val).strip() == "nan":
                return "Нет ответа"
            m = re.match(r"^(\d+)[\)\.]\s*(.*)", str(val).strip())
            if m:
                num = m.group(1)
                return mapping.get(num, str(val).strip())
            return str(val).strip()

        return series.apply(apply_map)
    else:
        return series.apply(
            lambda x: (
                str(x).strip()
                if pd.notna(x) and str(x).strip() != "nan"
                else "Нет ответа"
            )
        )


def clean_age(age_str):
    if pd.isna(age_str) or str(age_str).strip().lower() == "nan":
        return "Нет ответа"
    numbers = re.findall(r"\d+", str(age_str))
    if numbers:
        val = int(numbers[0])
        if val > 1900:
            return str(datetime.now().year - val)
        return str(val)
    return "Нет ответа"


def clean_dataframe(df):
    for col in df.columns:
        if is_system_column(col):
            continue
        elif "возраст" in col.lower() or "лет" in col.lower():
            df[col] = df[col].apply(clean_age)
        else:
            df[col] = unify_numbered_answers(df[col])

    df.rename(columns=lambda x: clean_column_name(x), inplace=True)
    return df


def get_column_groups(columns):
    groups = {}
    prefix_counts = {}
    for col in columns:
        if " / " in col:
            prefix = col.split(" / ")[0].strip()
            prefix_counts[prefix] = prefix_counts.get(prefix, 0) + 1

    for col in columns:
        if " / " in col:
            prefix = col.split(" / ")[0].strip()
            if prefix_counts[prefix] > 1:
                if prefix not in groups:
                    groups[prefix] = []
                groups[prefix].append(col)
            else:
                groups[col] = [col]
        else:
            groups[col] = [col]
    return groups


def _get_answer_counts(df, q_name, groups_cache, merged_sub=None):
    actual_cols = groups_cache[q_name]
    data = (
        df[actual_cols[0]] if len(actual_cols) == 1 else df[actual_cols].melt()["value"]
    )
    if merged_sub and merged_sub in groups_cache:
        sub_cols = groups_cache[merged_sub]
        sub_data = (
            df[sub_cols[0]] if len(sub_cols) == 1 else df[sub_cols].melt()["value"]
        )
        data = pd.concat([data, sub_data], ignore_index=True)
    data = data.dropna()
    data = data[data.astype(str).str.strip() != ""]
    data = data[data.astype(str).str.lower() != "nan"]
    data = data[data.astype(str) != "Нет ответа"]
    if data.empty:
        return {}
    counts = data.value_counts().reset_index()
    counts.columns = ["Ответ", "Количество"]
    counts["Ответ"] = counts["Ответ"].apply(clean_answer_text)
    counts = counts.groupby("Ответ", as_index=False)["Количество"].sum()
    return counts.set_index("Ответ")["Количество"].to_dict()


def generate_report_data(upload_dir, request_data):
    file_labels = request_data.file_labels
    file_colors = request_data.file_colors

    dfs = {}
    groups_cache = {}
    for clean_filename in file_labels.keys():
        filepath = os.path.join(upload_dir, clean_filename)
        if os.path.exists(filepath):
            df = pd.read_parquet(filepath)
            dfs[clean_filename] = df
            groups_cache[clean_filename] = get_column_groups(df.columns)

    results = []

    for cfg in request_data.configs:
        all_answers = set()
        file_counts = {}

        for f_name, q_name in cfg.file_mapping.items():
            if f_name not in dfs or q_name not in groups_cache[f_name]:
                continue
            counts = _get_answer_counts(
                dfs[f_name],
                q_name,
                groups_cache[f_name],
                merged_sub=cfg.merged_sub.get(f_name) if cfg.merged_sub else None,
            )
            if counts:
                file_counts[f_name] = counts
                all_answers.update(counts.keys())

        if not all_answers:
            continue

        data_dicts = []
        for ans in all_answers:
            row = {"answer": ans, "counts": {}, "included": True}
            row["counts"] = {
                f_name: file_counts.get(f_name, {}).get(ans, 0)
                for f_name in file_labels
            }
            row["_total"] = sum(row["counts"].values())
            data_dicts.append(row)

        data_dicts.sort(key=lambda x: x["_total"], reverse=True)

        section_data = None
        if hasattr(cfg, "section") and cfg.section:
            section_data = cfg.section

        results.append(
            {
                "col_name": cfg.column,
                "viz_type": cfg.viz_type,
                "data": data_dicts,
                "file_labels": file_labels,
                "file_colors": file_colors,
                "file_keys": list(file_labels.keys()),
                "section": section_data,
            }
        )

    return results
