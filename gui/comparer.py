import pandas as pd
from datetime import datetime

def normalize_text(val):
    if pd.isna(val):
        return ""
    return str(val).strip().lower()

def compare_excel_files(
    reference_path: str,
    target_path: str,
    save_result: bool = True,
    ignore_col_if_ref_value_empty: bool = False
) -> tuple[str, str]:
    """
    Compare two Excel files row-by-row and column-by-column based on the reference file.
    Normalize text before comparing and optionally skip columns if the reference cell is empty.

    Args:
        reference_path (str): Path to reference/template Excel file.
        target_path (str): Path to target Excel file.
        save_result (bool): Whether to save result back to target file.
        ignore_col_if_ref_value_empty (bool): If True, skip comparing a column at a row if
            the reference cell is empty for that row & column.

    Returns:
        Tuple[str, str]: (output_path, summary log text)
    """
    try:
        template_df = pd.read_excel(reference_path, dtype=str)
        scraped_df = pd.read_excel(target_path, dtype=str)

        compare_columns = list(template_df.columns)

        status_list = []
        status_msg_list = []
        start_time = datetime.now()

        for i in range(len(template_df)):
            if i >= len(scraped_df):
                status_list.append("MISMATCH")
                status_msg_list.append("Row missing in target file")
                continue

            diffs = []
            for col in compare_columns:
                val_ref_raw = template_df.iloc[i][col]
                val_scraped_raw = scraped_df.iloc[i][col]

                # Skip comparison if reference is empty ("" or NaN) or "#"
                if ignore_col_if_ref_value_empty and (pd.isna(val_ref_raw) or str(val_ref_raw).strip() in ["", "#"]):
                    continue

                val_ref = normalize_text(val_ref_raw)
                val_scraped = normalize_text(val_scraped_raw)

                if val_ref != val_scraped:
                    diffs.append(f"{col}: expected '{val_ref}', found '{val_scraped}'")

            if diffs:
                status_list.append("MISMATCH")
                status_msg_list.append(", ".join(diffs))
            else:
                status_list.append("MATCH")
                status_msg_list.append("OK")

        scraped_df['STATUS'] = status_list + ["MISSING ROW"] * (len(scraped_df) - len(status_list))
        scraped_df['STATUSMSG'] = status_msg_list + ["Row beyond reference"] * (len(scraped_df) - len(status_msg_list))
        scraped_df['STARTTIME'] = start_time
        scraped_df['ENDTIME'] = datetime.now()
        scraped_df['SCRAPINGSTATUS'] = "PROCESSED"

        output_path = target_path
        if save_result:
            scraped_df.to_excel(output_path, index=False)

        log_lines = [
            f"Row {idx + 1}: {status} - {msg}"
            for idx, (status, msg) in enumerate(zip(status_list, status_msg_list))
        ]
        log_text = "\n".join(log_lines)

        return output_path, log_text

    except Exception as e:
        raise RuntimeError(f"Comparison failed: {str(e)}")
