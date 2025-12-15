
import os
import argparse
import pandas as pd

def _normalize(s):
    return ''.join(str(s).strip().split()).lower()

def _find_column(df, target_set):
    for col in df.columns:
        if _normalize(col) in target_set:
            return col
    return None

def _prepare_keys(df, planned_col, patient_col):
    planned_series = pd.to_datetime(df[planned_col], errors='coerce')
    planned_key = planned_series.dt.date
    patient_key = df[patient_col].astype(str).str.strip()
    return planned_key, patient_key

def _valid_rows(df, planned_col, patient_col):
    planned_series = pd.to_datetime(df[planned_col], errors='coerce')
    patient_series = df[patient_col].astype(str).str.strip()
    mask = planned_series.notna() & patient_series.ne("")
    out_df = df.loc[mask].copy()
    return out_df

def _month_counts_from_df(df, planned_col):
    planned_series = pd.to_datetime(df[planned_col], errors='coerce')
    periods = planned_series.dt.to_period('M')
    counts_series = periods.value_counts(sort=False)
    counts_series = counts_series.sort_index()
    out = []
    for p, v in counts_series.items():
        label = str(p.month) + "/" + str(p.year)
        out.append((label, int(v)))
    return out

def process_with_months(input_path, do_dedupe, output_path=None):
    if not os.path.isfile(input_path):
        raise FileNotFoundError("Input file not found")
    df = pd.read_excel(input_path, engine='openpyxl')
    planned_targets = {"plannedstart", "planned_start"}
    patient_targets = {"patient", "patientname", "patient_name"}
    planned_col = _find_column(df, planned_targets)
    patient_col = _find_column(df, patient_targets)
    if planned_col is None or patient_col is None:
        raise ValueError("Required columns not found: Planned Start and Patient")
    df = _valid_rows(df, planned_col, patient_col)
    pre_dedupe_rows = len(df)
    out_df = df
    unique_count = None
    if do_dedupe:
        planned_key, patient_key = _prepare_keys(df, planned_col, patient_col)
        df['_planned_key'] = planned_key
        df['_patient_key'] = patient_key
        unique_df = df.drop_duplicates(subset=['_planned_key', '_patient_key'], keep='first').copy()
        unique_df[planned_col] = unique_df['_planned_key']
        unique_df = unique_df.drop(columns=['_planned_key', '_patient_key'])
        out_df = unique_df
        unique_count = len(unique_df)
    month_counts = _month_counts_from_df(out_df, planned_col)
    total_rows = len(out_df)
    if output_path is None:
        base_dir = os.path.dirname(os.path.abspath(input_path))
        output_path = os.path.join(base_dir, "planned_result.xlsx")
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        out_df.to_excel(writer, index=False, sheet_name='Unique' if do_dedupe else 'All')
    return output_path, unique_count, month_counts, total_rows, pre_dedupe_rows

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="Path to the input .xlsx file")
    parser.add_argument("--output", dest="output", default=None)
    parser.add_argument("--dedupe", dest="dedupe", action='store_true')
    args = parser.parse_args()
    out_path, unique_count, month_counts, total_rows, pre_dedupe_rows = process_with_months(args.input, args.dedupe, args.output)
    print(out_path)
    if unique_count is not None:
        print(str(unique_count))
    print(str(total_rows))
    print(str(pre_dedupe_rows))
    for label, v in month_counts:
        print(label + ":" + str(v))

if __name__ == "__main__":
    main()
