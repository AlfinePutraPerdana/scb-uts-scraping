from multiprocessing import context
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
import time
import traceback
from core.workflow import Workflow
import re
workflows = Workflow.load_all_workflows("workflow.json")


def find_table_by_header_keywords(driver, selector, header_keywords, header_row=1):
    candidates = driver.find_elements(By.CSS_SELECTOR, selector)
    print(f"[header_match] Found {len(candidates)} candidates for {selector}")

    def get_header_texts(table):
        try:
            row = table.find_element(By.CSS_SELECTOR, f"tbody > tr:nth-child({header_row})")
            return [cell.text.strip().lower() for cell in row.find_elements(By.CSS_SELECTOR, "th, td")]
        except:
            return []

    for idx, table in enumerate(candidates):
        headers = get_header_texts(table)
        print(f"[header_match] Table {idx+1} headers: {headers}")
        if all(any(keyword.lower() in header for header in headers) for keyword in header_keywords):
            print(f"✅ [header_match] Matched table {idx+1} with header keywords")
            return table

    raise Exception("No table matched the header keywords")

def scoped_find(driver, full_selector, table_index=None, label_keywords=None):
    from selenium.webdriver.common.by import By

    if ">" not in full_selector:
        return driver.find_element(By.CSS_SELECTOR, full_selector)

    parts = full_selector.split(">")
    base_selector = parts[0].strip()
    child_selector = ">".join(parts[1:]).strip()

    candidates = driver.find_elements(By.CSS_SELECTOR, base_selector)
    print(f"[scoped_find] Found {len(candidates)} candidate tables for selector: {base_selector}")

    if not candidates:
        raise Exception(f"No elements found for base selector: {base_selector}")

    # 1. Use table_index if provided
    if table_index is not None and 0 <= table_index - 1 < len(candidates):
        print(f"[scoped_find] Using table_index={table_index}")
        scope = candidates[table_index - 1]
        try:
            return scope.find_element(By.CSS_SELECTOR, child_selector)
        except:
            raise Exception(f"Element not found inside table_index {table_index} for selector: {child_selector}")

    # 2. Label keyword search
    if label_keywords:
        print(f"[scoped_find] Searching using label keywords: {label_keywords}")
        for idx, scope in enumerate(candidates):
            try:
                full_text = scope.text.lower()
                print(f"[scoped_find] Table {idx + 1}: '{full_text[:100]}'...")  # preview
                if all(kw.lower() in full_text for kw in label_keywords):
                    print(f"[scoped_find] ➕ Match found at table {idx + 1}")
                    return scope.find_element(By.CSS_SELECTOR, child_selector)
            except Exception as e:
                print(f"[scoped_find] ❌ Error checking table {idx + 1}: {e}")
                continue

    # 3. Fallback
    print(f"[scoped_find] Fallback: trying all candidates")
    for i, scope in enumerate(candidates):
        try:
            return scope.find_element(By.CSS_SELECTOR, child_selector)
        except:
            continue

    raise Exception(f"❌ Element not found using selector: {full_selector}")


def parse_selector_key(selector_key):
    parts = selector_key.split('/')
    if len(parts) == 4:
        return parts[0], parts[1], parts[2], parts[3]
    elif len(parts) == 3:
        return parts[0], parts[1], parts[2], None
    elif len(parts) == 2:
        return parts[0], parts[1], None, None
    else:
        raise ValueError(f"Invalid selector_key format: {selector_key}")


def run_workflow(driver, workflow, context, presets, team, screen, tab, log_callback=None):
    def log(msg):
        print(f"[debug] log() called with msg: {msg}")
        if callable(log_callback):
            log_callback(msg)
        else:
            print(msg)

    log("🚀 Starting workflow execution")
    log(f"📋 Total steps: {len(workflow)} | Context: {context}")
    results = {}
    window_stack = [driver.current_window_handle]
    
    if "results" not in context:
        context["results"] = {}  # ✅ Add this line
    
    if "current_excel_row" not in context:
        context["current_excel_row"] = {}  # ✅ Add this line

    for index, step in enumerate(workflow):
        action = step.get("action")
        selector_key = step.get("selector_key")  # Optional
        params = step.get("params", {})

        log(f"\n➡️ Step {index+1}/{len(workflow)}: Action = {action}")

        # Define actions that require selector_key
        selector_required_actions = {
            "click", "input", "switch_to_frame", "wait_for",
            "extract", "extract_dynamic_table", "extract_dynamic_table2",
            "extract_table_atm", "extract_dynamic_table_flat", "extract_table_risk_code", "extract_table_general_information", "extract_table_general_sweep", "extract_table_general_others", "extract_table_general_linkages", "extract_table_general_infotype", "extract_table_general_block", "extract_table_general_charge", "extract_fallback", "extract_table_new_atm"
        }

        selector = None
        if action in selector_required_actions:
            if not selector_key:
                log(f"⚠️ Step {index+1}: Skipped — missing selector_key for '{action}'")
                continue

            parts = selector_key.split("/")
            if len(parts) != 4:
                log(f"⚠️ Step {index+1}: Skipped — invalid selector_key format: '{selector_key}'")
                continue
            
            team, screen, tab, field = key_parts = parts


            try:
                selector_data = presets[team][screen][tab][field]
                selector = selector_data["selector"]
                print(f"Resolved selector: {selector}")
            except KeyError:
                log(f"❌ Selector not found for {team}/{screen}/{tab}/{field}")
                print(f"❌ Selector not found for {team}/{screen}/{tab}/{field}")
                if action in ["click", "input", "switch_to_frame", "wait_for"]:
                    raise
                else:
                    log(f"⚠️ Step {index+1}: Skipped — selector not found.")
                    continue
        else:
            log(f"📝 Step {index+1}: Skipping selector resolution for action '{action}'")


        try:
            if action == "switch_to_frame":
                log("🔄 Switching to frame...")
                print("🔄 Switching to frame...")
                driver.switch_to.default_content()

                print("Available iframes before switching:")
                for iframe in driver.find_elements(By.TAG_NAME, "iframe"):
                    print(f" - id='{iframe.get_attribute('id')}' name='{iframe.get_attribute('name')}'")

                for attempt in range(1, 31):
                    try:
                        frame_element = WebDriverWait(driver, 2).until(
                            EC.presence_of_element_located((By.CSS_SELECTOR, selector))
                        )
                        driver.switch_to.frame(frame_element)
                        log(f"✅ Switched to iframe '{selector}' on attempt {attempt}")
                        print(f"✅ Switched to iframe '{selector}' on attempt {attempt}")
                        break
                    except Exception as e:
                        print(f"Attempt {attempt}: Frame not ready... ({e})")
                        time.sleep(1)
                else:
                    raise Exception(f"Failed to switch to iframe '{selector}' after 30 seconds")
            
            elif action == "switch_to_new_window":
                current_handle = driver.current_window_handle
                all_handles = driver.window_handles
                target_title = params.get("target_title")
                print("🪟 Available windows:")

                for handle in all_handles:
                    driver.switch_to.window(handle)
                    title = driver.title or "[No title]"
                    url = driver.current_url
                    print(f"  • Handle: {handle[:8]}... | Title: {title} | URL: {url}")

                found = False
                for handle in all_handles:
                    if handle != current_handle:
                        driver.switch_to.window(handle)
                        WebDriverWait(driver, 5).until(lambda d: d.title != "")
                        if not target_title or target_title in driver.title:
                            print(f"✅ Switched to window: {driver.title}")

                            # ⬇️ Tambahkan handle lama ke stack sebelum pindah
                            if "window_stack" not in context:
                                context["window_stack"] = []
                            context["window_stack"].append(current_handle)

                            found = True
                            break

                if not found:
                    raise Exception(f"❌ No new window matched title: '{target_title}'")

            
            elif action == "close_window_and_return":
                driver.close()

                window_stack = context.get("window_stack", [])
                if window_stack:
                    previous_window = window_stack.pop()
                    driver.switch_to.window(previous_window)
                    if log_callback:
                        log_callback(f"↩️ Returned to previous window: {driver.title}")
                else:
                    raise Exception("No previous window to return to.")


            elif action == "extract_dynamic_table":
                try:
                    tab_preset = presets[team][screen][tab][field]
                    base_selector = tab_preset.get("selector", "#datatable")
                    column_map = tab_preset.get("column_map", {})
                    header_keywords = tab_preset.get("header_keywords", [])
                    header_row_index = tab_preset.get("header_row", 1)
                    row_offset = tab_preset.get("row_offset", header_row_index + 1)
                    order_by = params.get("order_by")
                    raw_slices = params.get("slices", [])
                    field_slices = {}
                    used_row_ids = set()
                    order_by_used = False

                    if isinstance(raw_slices, list):
                        for entry in raw_slices:
                            field = entry.get("field")
                            slice_spec = entry.get("slice")
                            if field and isinstance(slice_spec, list) and len(slice_spec) == 2:
                                field_slices[field] = slice_spec
                    elif isinstance(raw_slices, dict):
                        field_slices = raw_slices

                    max_rows = params.get("max_rows", 10)
                    conditions = params.get("conditions", {})
                    target_columns = params.get("target_columns")
                    output_map = params.get("output_map", {})
                    max_matches = params.get("max_matches", 3)
                    suffix_spacing = params.get("suffix_spacing", {})

                    current_excel_row = context.setdefault("current_excel_row", {})

                    matched_table = find_table_by_header_keywords(
                        driver,
                        base_selector,
                        header_keywords,
                        header_row=header_row_index
                    )

                    header_row = matched_table.find_element(By.CSS_SELECTOR, f"tbody > tr:nth-child({header_row_index})")
                    header_cells = header_row.find_elements(By.CSS_SELECTOR, "td, th")
                    header_map = {cell.text.strip().lower(): idx for idx, cell in enumerate(header_cells)}
                    print("[debug] Parsed header_map:", header_map)

                    from core.utils import passes_all_conditions
                    valid_rows = []
                    for row_idx in range(row_offset, row_offset + max_rows):
                        try:
                            row = matched_table.find_element(By.CSS_SELECTOR, f"tbody > tr:nth-child({row_idx})")
                            tds = row.find_elements(By.CSS_SELECTOR, "td")
                            visible_tds = [td for td in tds if td.value_of_css_property("display") != "none"]

                            print(f"[debug] Row {row_idx}: {len(tds)} tds, {len(visible_tds)} visible")
                            for i, td in enumerate(visible_tds, start=1):
                                print(f"  [col {i}] '{td.text.strip()}'")

                            row_dict = {}
                            for field_key, map_value in column_map.items():
                                try:
                                    if isinstance(map_value, int):
                                        if map_value <= len(visible_tds):
                                            value = visible_tds[map_value - 1].text.strip()
                                        else:
                                            value = "-"
                                            print(f"⚠️ [extract] Index {map_value} out of range for field '{field_key}'")
                                    else:
                                        col_index = header_map.get(map_value.lower())
                                        if col_index is not None and col_index < len(visible_tds):
                                            value = visible_tds[col_index].text.strip()
                                        else:
                                            value = "-"
                                            print(f"⚠️ [extract] Missing header '{map_value}' for field '{field_key}'")

                                    # Before slicing
                                    original_value = value
                                    slice_spec = field_slices.get(field_key)
                                    if isinstance(slice_spec, list) and len(slice_spec) == 2 and isinstance(value, str):
                                        value = value[slice_spec[0]:slice_spec[1]]
                                        print(f"[slice] {field_key}: '{original_value}' -> '{value}' (slice {slice_spec})")
                                    row_dict[field_key] = value

                                    if field_key.lower() == "attention_party":
                                        print(f"[debug] Attention_Party @ row {row_idx}: '{value}'")

                                except Exception as e:
                                    print(f"⚠️ Error extracting field '{field_key}': {e}")
                                    row_dict[field_key] = "-"
                                    continue

                            if conditions and not passes_all_conditions(row_dict, conditions, context.get("results", {})):
                                print(f"[filter] Skipping row {row_idx}: doesn't match conditions")
                                continue

                            if target_columns:
                                row_dict = {k: v for k, v in row_dict.items() if k in target_columns}

                            valid_rows.append(row_dict)

                        except Exception as e:
                            print(f"⚠️ Error on row {row_idx}: {e}")
                            break

                    extracted = 0
                    if isinstance(order_by, dict):
                        sort_field = order_by.get("field")
                        sort_values = order_by.get("values", [])
                        sort_index = {v: i for i, v in enumerate(sort_values)}
                        slice_spec = field_slices.get(sort_field)
                        used_row_ids = set()
                        print(f"[order_by] Writing in fixed order: {sort_values}")

                        for target_index, expected_value in enumerate(sort_values):
                            matched_row = None
                            for row_idx, row in enumerate(valid_rows):
                                if row_idx in used_row_ids:
                                    continue
                                actual_value = row.get(sort_field, "")
                                # DO NOT SLICE HERE! Assume already sliced in valid_rows.
                                if actual_value == expected_value:
                                    matched_row = row
                                    used_row_ids.add(row_idx)
                                    break

                            for write_key in target_columns or (matched_row.keys() if matched_row else []):
                                raw_output_key = output_map.get(write_key, write_key)
                                if "{n}" in raw_output_key:
                                    excel_key = raw_output_key.replace("{n}", str(target_index + 1))
                                else:
                                    spacing = suffix_spacing.get(raw_output_key, False)
                                    suffix = "" if target_index == 0 else (f" {target_index + 1}" if spacing else f"{target_index + 1}")
                                    excel_key = f"{raw_output_key}{suffix}"
                                value = matched_row.get(write_key, "") if matched_row else ""
                                if isinstance(value, str):
                                    value = value.strip()
                                context["current_excel_row"][excel_key] = value
                            
                            print(f"[write] {expected_value} → {matched_row if matched_row else '❌ not found'}")
                            extracted += 1
                            if extracted >= max_matches:
                                print(f"[info] Reached max_matches ({max_matches}); stopping.")
                                break

                    else:
                        for match_idx, row_dict in enumerate(valid_rows):
                            if target_columns:
                                write_keys = [k for k in target_columns if k in row_dict]
                            else:
                                write_keys = list(row_dict.keys())
                            for field_key in write_keys:
                                value = row_dict[field_key]
                                slice_spec = field_slices.get(field_key)
                                if isinstance(slice_spec, list) and len(slice_spec) == 2:
                                    value = value[slice_spec[0]:slice_spec[1]]
                                base_excel_key = output_map.get(field_key, field_key)
                                if "{n}" in base_excel_key:
                                    excel_key = base_excel_key.replace("{n}", str(match_idx + 1))
                                else:
                                    space_required = suffix_spacing.get(base_excel_key, True)
                                    suffix = "" if match_idx == 0 else f"{' ' if space_required else ''}{match_idx + 1}"
                                    excel_key = f"{base_excel_key}{suffix}"
                                current_excel_row[excel_key] = value
                                extracted += 1
                            if match_idx + 1 >= max_matches:
                                print(f"[info] Reached max_matches ({max_matches}); stopping.")
                                break
                    
                    
                    context["last_dynamic_count"] = extracted
                    context["last_output_base"] = output_map

                    print(f"[after extract_dynamic_table] current_excel_row keys: {list(current_excel_row.keys())}")

                    log(f"📄 Step {index+1}: Extracted {extracted} values from dynamic table.")
                    for key, val in current_excel_row.items():
                        print(f"  - {key}: {val}")
                    print("🧾 Final current_excel_row contents:")
                    for k, v in sorted(current_excel_row.items()):
                        print(f"  - {k}: {v}")

                except Exception as e:
                    log(f"❌ Step {index+1}: Failed to extract dynamic table — {e}")
                    print(f"❌ Step {index+1}: Failed to extract dynamic table — {traceback.format_exc()}")
                    continue
                
            elif action == "extract_dynamic_table_general":
                try:
                    tab_preset = presets[team][screen][tab][field]
                    base_selector = tab_preset.get("selector", "#datatable")
                    column_map = tab_preset.get("column_map", {})
                    header_keywords = tab_preset.get("header_keywords", [])
                    header_row_index = tab_preset.get("header_row", 1)
                    row_offset = tab_preset.get("row_offset", header_row_index + 1)
                    max_rows = params.get("max_rows", 10)
                    conditions = params.get("conditions", {})

                    # Simpan hasil table di list khusus
                    if "dynamic_table_results" not in context:
                        context["dynamic_table_results"] = []
                    if "dynamic_table_results" not in results:
                        results["dynamic_table_results"] = []

                    matched_table = find_table_by_header_keywords(
                        driver,
                        base_selector,
                        header_keywords,
                        header_row=header_row_index
                    )

                    # Buat header_map
                    header_row = matched_table.find_element(By.CSS_SELECTOR, f"tbody > tr:nth-child({header_row_index})")
                    header_cells = header_row.find_elements(By.CSS_SELECTOR, "td, th")
                    header_map = {cell.text.strip().lower(): idx for idx, cell in enumerate(header_cells)}
                    print("[debug] Parsed header_map:", header_map)

                    extracted = 0
                    for row_idx in range(row_offset, row_offset + max_rows):
                        try:
                            row = matched_table.find_element(By.CSS_SELECTOR, f"tbody > tr:nth-child({row_idx})")
                            tds = row.find_elements(By.CSS_SELECTOR, "td")
                            visible_tds = [td for td in tds if td.value_of_css_property("display") != "none"]

                            row_dict = {}
                            for field_key, map_value in column_map.items():
                                try:
                                    if isinstance(map_value, int):
                                        if 0 < map_value <= len(visible_tds):
                                            value = visible_tds[map_value - 1].text.strip()
                                        else:
                                            value = "-"
                                    else:
                                        col_index = header_map.get(map_value.lower())
                                        if col_index is not None and col_index < len(visible_tds):
                                            value = visible_tds[col_index].text.strip()
                                        else:
                                            value = "-"
                                    row_dict[field_key] = value
                                except Exception:
                                    row_dict[field_key] = "-"
                                    continue

                            # Filter by conditions
                            if conditions:
                                match = all(row_dict.get(k, "").strip() == v for k, v in conditions.items())
                                if not match:
                                    continue

                            # Simpan ke context dan results
                            context["dynamic_table_results"].append(row_dict)
                            results["dynamic_table_results"].append(row_dict)
                            extracted += 1

                        except Exception as e:
                            print(f"⚠️ Error on row {row_idx}: {e}")
                            break

                    log(f"📄 Step {index+1}: Extracted {extracted} rows from dynamic table.")

                except Exception as e:
                    log(f"❌ Step {index+1}: Failed to extract dynamic table — {e}")
                    print(f"❌ Step {index+1}: Failed to extract dynamic table — {traceback.format_exc()}")
                    continue


            elif action == "extract_dynamic_table2":
                try:
                    tab_preset = presets[team][screen][tab][field]
                    base_selector = tab_preset.get("selector", "#datatable")
                    selector_template = f"{base_selector} > tbody > tr:nth-child({{row}}) > td:nth-child({{col}})"
                    column_map = tab_preset.get("column_map", {})

                    # Determine which columns to extract
                    target_columns = params.get("target_columns", None)
                    if target_columns:
                        active_map = {k: v for k, v in column_map.items() if k in target_columns}
                    else:
                        active_map = column_map

                    rows = driver.find_elements(By.CSS_SELECTOR, f"{base_selector} > tbody > tr")
                    print(f"[debug] Found {len(rows)} rows in dynamic table2")

                    extracted = 3  # your original skip logic
                    extracted_rows = []
                    for index, row in enumerate(rows):
                        if row.text.strip() == "":
                            continue
                        if index == 0:
                            print(f"[debug] Skipping header row: {row.text.strip()}")
                            continue

                        print(f"[debug] Processing row {index}: {row.text.strip()}")
                        row_data = {}
                        for field_key, col_index in active_map.items():
                            value = ""
                            try:
                                selector = selector_template.format(row=extracted, col=col_index)
                                cell = driver.find_element(By.CSS_SELECTOR, selector)
                                value = cell.text.strip()

                                if not value:
                                    # fallback: look for first meaningful child element
                                    try:
                                        child = cell.find_element(By.CSS_SELECTOR, "a, span, nobr, *")
                                        value = child.text.strip()
                                    except:
                                        pass
                            except:
                                value = ""

                            row_data[field_key] = value
                        extracted_rows.append(row_data)
                        extracted += 2  # skip to next logical row

                    results[field] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} row(s) to '{field}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} row(s) to '{field}'")
                except Exception as e:
                    print(f"⚠️ Dynamic table2 error:\n{traceback.format_exc()}")
                    log(f"⚠️ Step {index+1}: Skipped dynamic table2 — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Skipped dynamic table2 — {str(e).splitlines()[0]}")
                    continue



                
            elif action == "extract_table_atm":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector
                    
                    log(f"🔍 [ATM] Using selector: {table_selector}")
                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    log(f"📑 Found {len(rows)} row(s) in table")

                    # 🔍 Coba deteksi apakah baris pertama adalah header
                    if rows:
                        first_cells = rows[0].find_elements(By.CSS_SELECTOR, "td")
                        if any("card" in cell.text.lower() or "account" in cell.text.lower() for cell in first_cells):
                            log(f"⚪️ [Row 1] Skipped: kemungkinan baris header.")
                            rows = rows[1:]  # skip baris header

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        # Skip empty rows
                        if not cells or all(c.text.strip() == "" for c in cells):
                            log(f"⚪️ [Row {row_index}] Skipped: empty or blank")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        row_data = {}
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")
                                
                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} ATM row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} ATM row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract ATM table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract ATM table:\n{traceback.format_exc()}")
                    continue

            elif action == "extract_table_new_atm":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector

                    account_no = context.get("Account No", "")
                    ccy = context.get("CCY", "")

                    print(f"🔍 [ATM] Using: {account_no}, CCY={ccy}")
                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    print(f"📑 Found {len(rows)} row(s) in table")

                    # Daftar kata kunci header
                    header_keywords = {
                        "account no", "ccy", "card no", "card seq no", "card name",
                        "card status", "record status", "withdrawal limit", "card expiry date",
                        "card type", "account short name", "maker id", "maker date", "maker time",
                        "relationship no", "card activation date", "card channel id", "card activation time",
                        "overseas activation/deactivation status", "overseas activation date",
                        "overseas activation time", "overseas activation channel id",
                        "overseas deactivation date", "overseas deactivation time",
                        "overseas deactivation channel id", "next charge apply date",
                        "renewal expiry date", "renewal card status", "re issue date",
                        "card transacted", "overseas startdate", "overseas enddate",
                        "related relationship no"
                    }

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        texts = [c.text.strip().lower() for c in cells if c.text.strip()]

                        if not texts:
                            log(f"⚪️ [Row {row_index}] Skipped: empty or blank")
                            continue

                        # --- Skip jika baris terdeteksi sebagai header ---
                        header_match_count = sum(1 for t in texts if any(k in t for k in header_keywords))
                        if header_match_count >= max(1, int(len(texts) * 0.7)):
                            log(f"⚪️ [Row {row_index}] Skipped: detected header row ({texts})")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        row_data = {
                            "AccountNo": account_no,
                            "CCY": ccy
                        }
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")

                        # --- Skip kalau row_data masih persis sama dengan header ---
                        header_like = sum(1 for k, v in row_data.items() if v.strip().lower() in header_keywords)
                        if header_like >= max(1, int(len(row_data) * 0.7)):
                            log(f"⚪️ [Row {row_index}] Skipped: looks like header row_data ({row_data})")
                            continue

                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} ATM row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} ATM row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract ATM table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract ATM table:\n{traceback.format_exc()}")
                    continue

            elif action == "extract_table_general_information":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector

                    # Ambil AccountNo & CCY dari context
                    account_no = context.get("Account No", "")
                    ccy = context.get("CCY", "")

                    log(f"🔍 [General Info] Using AccountNo={account_no}, CCY={ccy}")
                    log(f"🔍 [General Info] Selector: {table_selector}")

                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    log(f"📑 Found {len(rows)} row(s) in table")

                    # Helper untuk deteksi header row
                    def is_header_row(cells, keywords):
                        texts = [c.text.strip().lower() for c in cells if c.text.strip()]
                        return all(any(k.lower() in t for k in keywords) for t in texts) if texts else False

                    # Header keywords (gabungan untuk Block, Charge, Info, dsb.)
                    header_keywords = {
                        # Common
                        "account", "currency", "code", "flag", "balance", "status",
                        # Linkages
                        "relationshipid", "relationshipname", "linktype",
                        # InfoType
                        "info code", "seq no",
                        # Block tab
                        "block no.", "block type", "lien", "narrative", "expiry date",
                        "pending amount", "instruction no.", "sweep no.",
                        # Charge tab
                        "nominated currency", "charge group", "charge code",
                        "waiver reason", "percentage", "discount"
                    }

                    # Keywords untuk deteksi baris invalid (bukan data)
                    na_keywords = {
                        "no account to account linkage details available",
                        "no account to deals linkage details available",
                        "no cbod to deals linkage details available",
                        "no account links available",
                        "no account to loan linkage details available",
                        "not available"
                    }

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        texts = [c.text.strip().lower() for c in cells if c.text.strip()]

                        # Skip jika kosong total
                        if not texts:
                            log(f"⚪️ [Row {row_index}] Skipped: empty or blank")
                            continue

                        # Skip header row (detected by keywords)
                        if is_header_row(cells, header_keywords):
                            log(f"⚪️ [Row {row_index}] Skipped: detected header row ({texts})")
                            continue

                        # Skip invalid "Not Available" row
                        if any(na in " ".join(texts) for na in na_keywords):
                            log(f"⚪️ [Row {row_index}] Skipped: 'Not Available' row ({texts})")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        # Awal data: sisipkan AccountNo & CCY
                        row_data = {
                            "AccountNo": account_no,
                            "CCY": ccy
                        }

                        # Extract per column mapping
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")

                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract General Information table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract General Information table:\n{traceback.format_exc()}")
                    continue

            elif action == "extract_table_general_block":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector

                    account_no = context.get("Account No", "")
                    ccy = context.get("CCY", "")

                    print(f"🔍 [General Block] Using: {account_no}, CCY={ccy}")
                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    print(f"📑 Found {len(rows)} row(s) in table")

                    # Daftar kata kunci header Block
                    header_keywords = {
                        "block no", "date", "block type", "dr account", "cr account",
                        "block amount", "pending amount", "dr narrative", "cr narrative",
                        "instruction no", "sweep no", "rel channel id", "expiry date",
                        "post after date", "status", "lien/block source", "maker id",
                        "ebbs account reference", "lien refernce account"
                    }

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        texts = [c.text.strip().lower() for c in cells if c.text.strip()]

                        if not texts:
                            log(f"⚪️ [Row {row_index}] Skipped: empty or blank")
                            continue

                        # 🔍 Skip jika semua cell cocok header
                        if texts and all(any(k in t for k in header_keywords) for t in texts):
                            log(f"⚪️ [Row {row_index}] Skipped: detected header row ({texts})")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        row_data = {
                            "AccountNo": account_no,
                            "CCY": ccy
                        }
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")

                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} Block row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} Block row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract Block table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract Block table:\n{traceback.format_exc()}")
                    continue


            elif action == "extract_table_general_charge":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector

                    # Ambil AccountNo & CCY dari context (sesuai key di sheet utama)
                    account_no = context.get("Account No", "")
                    ccy = context.get("CCY", "")
                    
                    print(f"🔍 [General Others] Using: {account_no}")
                    print(f"🔍 [General Others] Using selector: {table_selector}")
                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    print(f"📑 Found {len(rows)} row(s) in table")

                    # 🔍 Skip baris pertama kalau terdeteksi header
                    if rows:
                        first_cells = rows[0].find_elements(By.CSS_SELECTOR, "td")
                        first_texts = [c.text.strip().lower() for c in first_cells]
                        if any("rls account number" in t or "selfrecontype" in t or "m1 flag" in t for t in first_texts):
                            print(f"⚪️ [Row 1] Skipped: kemungkinan baris header.")
                            rows = rows[1:]

                    # Semua keyword diturunkan ke lowercase
                    header_keywords = {
                        "nominated currency", "nominated account no", "charge group", "charge code / charge group code",
                        "charge waived", "charge offset", "discount flag", "percentage",
                        "waiver reason", "review date", "no of discounts", "discounts available"
                    }

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        texts = [c.text.strip().lower() for c in cells if c.text.strip()]

                        # Skip second header jika semua cell ada di daftar header_keywords
                        if texts and all(t in header_keywords for t in texts):
                            log(f"⚪️ [Row {row_index}] Skipped: detected as second header ({texts})")
                            continue

                        # Skip empty rows
                        if not cells or all(c.text.strip() == "" for c in cells):
                            log(f"⚪️ [Row {row_index}] Skipped: empty or blank")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        # Awal data: langsung sisipkan AccountNo & CCY
                        row_data = {
                            "AccountNo": account_no,
                            "CCY": ccy
                        }
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")
                                
                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract General Others table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract General Others table:\n{traceback.format_exc()}")
                    continue


            elif action == "extract_table_general_others":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector

                    # Ambil AccountNo & CCY dari context (sesuai key di sheet utama)
                    account_no = context.get("Account No", "")
                    ccy = context.get("CCY", "")
                    
                    print(f"🔍 [General Others] Using: {account_no}")
                    print(f"🔍 [General Others] Using selector: {table_selector}")
                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    print(f"📑 Found {len(rows)} row(s) in table")

                    # 🔍 Skip baris pertama kalau terdeteksi header
                    if rows:
                        first_cells = rows[0].find_elements(By.CSS_SELECTOR, "td")
                        first_texts = [c.text.strip().lower() for c in first_cells]
                        if any("rls account number" in t or "selfrecontype" in t or "m1 flag" in t for t in first_texts):
                            print(f"⚪️ [Row 1] Skipped: kemungkinan baris header.")
                            rows = rows[1:]

                    # Semua keyword diturunkan ke lowercase
                    header_keywords = {
                        "rls account number", "selfrecontype", "m1 flag", "secret / cic",
                        "phone banking flag", "dsr refereralid", "dsr sourcingid", "dsr closingid",
                        "limit no", "limit master no", "suspend impairment", "suspend provision",
                        "suspend chargeoff", "daue limitno", "daue limit masterno",
                        "intraday limitno", "intraday limit masterno"
                    }

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        texts = [c.text.strip().lower() for c in cells if c.text.strip()]

                        # Skip second header jika semua cell ada di daftar header_keywords
                        if texts and all(t in header_keywords for t in texts):
                            log(f"⚪️ [Row {row_index}] Skipped: detected as second header ({texts})")
                            continue

                        # Skip empty rows
                        if not cells or all(c.text.strip() == "" for c in cells):
                            log(f"⚪️ [Row {row_index}] Skipped: empty or blank")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        # Awal data: langsung sisipkan AccountNo & CCY
                        row_data = {
                            "AccountNo": account_no,
                            "CCY": ccy
                        }
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")
                                
                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract General Others table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract General Others table:\n{traceback.format_exc()}")
                    continue

            elif action == "extract_table_general_infotype":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector

                    # Ambil AccountNo & CCY dari context
                    account_no = context.get("Account No", "")
                    ccy = context.get("CCY", "")
                    
                    log(f"🔍 [General Info] Using AccountNo={account_no}, CCY={ccy}")
                    log(f"🔍 [General Info] Selector: {table_selector}")

                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    log(f"📑 Found {len(rows)} row(s) in table")

                    # Helper untuk deteksi header row
                    def is_header_row(cells, keywords):
                        texts = [c.text.strip().lower() for c in cells if c.text.strip()]
                        return all(any(k.lower() in t for k in keywords) for t in texts) if texts else False

                    # Keyword untuk header
                    header_keywords = {
                        "Info Code", "Info Code Description", "Info Code Lang Description",
                        "Info Details Code", "Info Detail Code Description", "Seq No"
                    }

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")

                        # Skip empty row
                        if not cells or all(c.text.strip() == "" for c in cells):
                            log(f"⚪️ [Row {row_index}] Skipped: empty or blank")
                            continue

                        # Skip header row (baik di atas maupun yang duplikat di tengah tabel)
                        if is_header_row(cells, header_keywords):
                            log(f"⚪️ [Row {row_index}] Skipped: detected as header row")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        # Awal data: sisipkan AccountNo & CCY
                        row_data = {
                            "AccountNo": account_no,
                            "CCY": ccy
                        }

                        # Extract per column mapping
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")

                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} InfoType row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} InfoType row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract InfoType table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract InfoType table:\n{traceback.format_exc()}")
                    continue


            elif action == "extract_table_general_linkages":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector

                    # Ambil AccountNo & CCY dari context (sheet utama)
                    account_no = context.get("Account No", "")
                    ccy = context.get("CCY", "")

                    log(f"🔍 [General Info] Using AccountNo={account_no}, CCY={ccy}")
                    log(f"🔍 [General Info] Using selector: {table_selector}")

                    # Ambil semua baris dalam tabel
                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    log(f"📑 Found {len(rows)} row(s) in table")

                    # Helper untuk cek apakah row adalah header utama
                    def is_main_header(cells):
                        texts = [c.text.strip().lower() for c in cells]
                        return any(
                            "relationshipid" in t or "relationshipname" in t or "linktype" in t
                            for t in texts
                        )

                    # Keyword untuk "fake header" (baris info kosong)
                    header_keywords = [
                        "no account to account linkage details available",
                        "no account to deals linkage details available",
                        "no cbod to deals linkage details available",
                        "no account links available",
                        "no account to loan linkage details available",
                    ]

                    def is_fake_header(cells):
                        texts = [c.text.strip().lower() for c in cells if c.text.strip()]
                        if not texts:
                            return False
                        # Kalau setiap cell punya keyword "No ..." → fake header
                        return any(any(keyword in t for keyword in header_keywords) for t in texts)

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")

                        # Skip: header utama
                        if is_main_header(cells):
                            log(f"⚪️ [Row {row_index}] Skipped: detected as main header")
                            continue

                        # Skip: fake header (No Account To ...)
                        if is_fake_header(cells):
                            log(f"⚪️ [Row {row_index}] Skipped: detected as fake header")
                            continue

                        # Skip: empty row
                        if not cells or all(c.text.strip() == "" for c in cells):
                            log(f"⚪️ [Row {row_index}] Skipped: empty row")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        # Awal data: sisipkan AccountNo & CCY
                        row_data = {
                            "AccountNo": account_no,
                            "CCY": ccy
                        }

                        # Extract sesuai mapping kolom
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")

                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} linkage row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} linkage row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract Linkages table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract Linkages table:\n{traceback.format_exc()}")
                    continue


            elif action == "extract_table_general_sweep":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector

                    # Ambil AccountNo & CCY dari context (sesuai key di sheet utama)
                    account_no = context.get("Account No", "")
                    ccy = context.get("CCY", "")
                    
                    print(f"🔍 [General Info] Using: {account_no}")
                    print(f"🔍 [General Info] Using selector: {table_selector}")
                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    print(f"📑 Found {len(rows)} row(s) in table")

                    # 🔍 Coba deteksi apakah baris pertama adalah header
                    # 🔍 Skip baris header pertama (main header) kalau cocok
                    if rows:
                        first_cells = rows[0].find_elements(By.CSS_SELECTOR, "td")
                        if any("card" in cell.text.lower() or "account" in cell.text.lower() or "currency" in cell.text.lower() for cell in first_cells):
                            print(f"⚪️ [Row 1] Skipped: kemungkinan baris header.")
                            rows = rows[1:]  # skip baris header

                    # Keyword untuk deteksi "second header"
                    header_keywords = {"currency", "account no", "account", "balance", "flag"}

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        texts = [c.text.strip().lower() for c in cells]

                        # Skip second header jika semua cell cocok keyword
                        if cells and all(any(k in t for k in header_keywords) for t in texts if t):
                            log(f"⚪️ [Row {row_index}] Skipped: detected as second header ({texts})")
                            continue

                        # Skip empty rows
                        if not cells or all(c.text.strip() == "" for c in cells):
                            log(f"⚪️ [Row {row_index}] Skipped: empty or blank")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        # Awal data: langsung sisipkan AccountNo & CCY
                        row_data = {
                            "AccountNo": account_no,
                            "CCY": ccy
                        }
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")
                                
                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} ATM row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} ATM row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract ATM table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract ATM table:\n{traceback.format_exc()}")
                    continue
            
            
            elif action == "extract_table_risk_code":
                try:
                    column_map = params.get("columns", {})
                    target_column = params.get("target_column", field)
                    table_selector = selector

                    # Ambil AccountNo & CCY dari context (sesuai key di sheet utama)
                    account_no = context.get("Account No", "")
                    ccy = context.get("CCY", "")
                    
                    log(f"🔍 [ATM] Using selector: {table_selector}")
                    rows = driver.find_elements(By.CSS_SELECTOR, f"{table_selector} > tbody > tr")
                    log(f"📑 Found {len(rows)} row(s) in table")

                    # 🔍 Coba deteksi apakah baris pertama adalah header
                    if rows:
                        first_cells = rows[0].find_elements(By.CSS_SELECTOR, "td")
                        if any("risk" in cell.text.lower() or "description" in cell.text.lower() for cell in first_cells):
                            log(f"⚪️ [Row 1] Skipped: kemungkinan baris header.")
                            rows = rows[1:]  # skip baris header

                    extracted_rows = []

                    for row_index, row in enumerate(rows, start=1):
                        cells = row.find_elements(By.CSS_SELECTOR, "td")
                        # Skip empty rows
                        if not cells or all(c.text.strip() == "" for c in cells):
                            log(f"⚪️ [Row {row_index}] Skipped: empty or blank")
                            continue

                        log(f"📥 [Row {row_index}] {len(cells)} cells detected")

                        # Awal data: langsung sisipkan AccountNo & CCY
                        row_data = {
                            "AccountNo": account_no,
                            "CCY": ccy
                        }

                        # Ambil kolom sesuai mapping
                        for key, col_css in column_map.items():
                            try:
                                cell = row.find_element(By.CSS_SELECTOR, col_css)
                                value = cell.text.strip()
                                row_data[key] = value
                                log(f"📥 [Row {row_index}] Extracted '{key}': '{value}'")
                            except:
                                row_data[key] = ""
                                log(f"⚠️ [Row {row_index}] Failed to extract '{key}' using '{col_css}'")
                                
                        extracted_rows.append(row_data)

                    results[target_column] = extracted_rows
                    log(f"📊 Step {index+1}: Extracted {len(extracted_rows)} ATM row(s) to '{target_column}'")
                    print(f"📊 Step {index+1}: Extracted {len(extracted_rows)} ATM row(s) to '{target_column}'")

                except Exception as e:
                    log(f"⚠️ Step {index+1}: Failed to extract ATM table — {str(e).splitlines()[0]}")
                    print(f"⚠️ Step {index+1}: Failed to extract ATM table:\n{traceback.format_exc()}")
                    continue

            elif action == "click":
                log("🖱 Clicking...")
                driver.find_element(By.CSS_SELECTOR, selector).click()
                log(f"🖱 Step {index+1}: Clicked.")
                print(f"🖱 Step {index+1}: Clicked.")

            elif action == "input":
                value = context.get(step.get("value_from") or params.get("context_key"), "")
                log(f"⌨️ Inputting: {value}")
                print(f"⌨️ Inputting: {value}")
                table_index = selector_data.get("table_index")
                el = scoped_find(driver, selector, table_index=table_index)
                el.clear()
                el.send_keys(str(value))
                log(f"⌨️ Step {index+1}: Input '{value}'.")
                print(f"⌨️ Step {index+1}: Input '{value}'.")

            elif action == "wait_for":
                log("⏳ Waiting for element...")
                optional = step.get("params", {}).get("optional", False)

                try:
                    wait_selector = None
                    # Dynamic table: wait for row_selector instead
                    if isinstance(selector_data, dict) and selector_data.get("type") == "dynamic_table":
                        wait_selector = selector_data.get("row_selector")
                        log(f"⏳ Waiting for first row of dynamic table: {wait_selector}")
                    else:
                        wait_selector = selector_data.get("selector")

                    WebDriverWait(driver, 10).until(
                        EC.presence_of_element_located((By.CSS_SELECTOR, wait_selector))
                    )
                    log(f"⏳ Step {index+1}: Element appeared.")

                except TimeoutException:
                    if optional:
                        log(f"⚠️ Step {index+1}: Optional element not found, continuing.")
                    else:
                        raise

            elif action == "extract":
                try:
                    
                    from core.utils import passes_all_conditions

                    # ✅ Evaluate conditions (if any)
                    print("[debug] context['results']:", context.get("results", {}))
                    print("[debug] context['current_excel_row']:", context.get("current_excel_row", {}))
                    conditions = step.get("conditions", []) or params.get("conditions", [])
                    print(f"[debug] Step {index+1}: conditions={conditions}")  # ← ADD THIS
                    if conditions:
                        merged_context = {
                            **context.get("results", {}),
                            **context.get("current_excel_row", {})
                        }

                        print(f"[debug] Step {index+1}: Merged context for condition check: {merged_context}")

                        if not passes_all_conditions(merged_context, conditions, merged_context):
                            print(f"[extract] Skipped due to failing conditions: {conditions}")
                            log(f"⏭️ Step {index+1}: Skipped due to condition not met.")
                            continue

                    table_index = selector_data.get("table_index")
                    label_keywords = selector_data.get("label_keywords")  # optional
                    print(f"[debug] Step {index+1}: table_index={table_index}, label_keywords={label_keywords}")
                    
                    
                    el = scoped_find(driver, selector, table_index=table_index, label_keywords=label_keywords)
                    text = el.text.strip()

                    # ✅ Optional slice before saving
                    slice_spec = params.get("slice")
                    if isinstance(slice_spec, list) and len(slice_spec) == 2:
                        text = text[slice_spec[0]:slice_spec[1]]


                    # ✅ Optional slice before saving
                    slice_spec = params.get("slice")
                    if isinstance(slice_spec, list) and len(slice_spec) == 2:
                        text = text[slice_spec[0]:slice_spec[1]]
                    
                    # ✅ Optional remove characters
                    remove_chars = params.get("remove_chars")  # e.g., ","
                    if remove_chars:
                        for ch in remove_chars:
                            text = text.replace(ch, "")

                    target_column = step.get("as") or params.get("target_column") or field
                    results[target_column] = text
                    
                    # ✅ Make extracted value immediately visible to later steps
                    if "current_excel_row" not in context:
                        context["current_excel_row"] = {}
                    context["current_excel_row"][target_column] = text

                    
                    # ✅ Make extracted value immediately visible to later steps
                    if "current_excel_row" not in context:
                        context["current_excel_row"] = {}
                    context["current_excel_row"][target_column] = text

                    log(f"📤 Step {index+1}: Extracted {target_column} = '{text}'")
                    print(f"📤 Step {index+1}: Extracted {target_column} = '{text}'")

                except Exception as e:
                    print(f"⚠️ Extract skipped for {field}:\n{traceback.format_exc()}")
                    log(f"⚠️ Step {index+1}: Extract skipped for {field} — {str(e).splitlines()[0]}")
                    continue
            
            elif action == "extract_fallback":
                try:
                    from core.utils import passes_all_conditions

                    # ✅ Evaluate conditions (if any)
                    conditions = step.get("conditions", []) or params.get("conditions", [])
                    if conditions:
                        merged_context = {
                            **context.get("results", {}),
                            **context.get("current_excel_row", {})
                        }
                        if not passes_all_conditions(merged_context, conditions, merged_context):
                            log(f"⏭️ Step {index+1}: Skipped due to condition not met.")
                            continue

                    table_index = selector_data.get("table_index")
                    label_keywords = selector_data.get("label_keywords")
                    selector_value = selector  # bisa string atau list

                    el, text = None, ""

                    # 🔄 Support multiple selectors with try/except fallback
                    if isinstance(selector_value, list):
                        for sel in selector_value:
                            try:
                                el = scoped_find(driver, sel, table_index=table_index, label_keywords=label_keywords)
                                if el and el.text.strip():
                                    text = el.text.strip()
                                    print(f"[debug] Step {index+1}: Found element with selector {sel} → '{text}'")
                                    break
                            except Exception as inner_e:
                                print(f"[debug] Step {index+1}: Selector {sel} failed: {inner_e}")
                                continue
                    else:
                        try:
                            el = scoped_find(driver, selector_value, table_index=table_index, label_keywords=label_keywords)
                            if el:
                                text = el.text.strip()
                        except Exception as inner_e:
                            print(f"[debug] Step {index+1}: Single selector {selector_value} failed: {inner_e}")

                    if not text:
                        log(f"⚠️ Step {index+1}: No element found for {selector_value}")
                        continue

                    # ✅ Optional slice before saving
                    slice_spec = params.get("slice")
                    if isinstance(slice_spec, list) and len(slice_spec) == 2:
                        text = text[slice_spec[0]:slice_spec[1]]

                    # ✅ Optional remove characters
                    remove_chars = params.get("remove_chars")
                    if remove_chars:
                        for ch in remove_chars:
                            text = text.replace(ch, "")

                    target_column = step.get("as") or params.get("target_column") or field
                    results[target_column] = text

                    if "current_excel_row" not in context:
                        context["current_excel_row"] = {}
                    context["current_excel_row"][target_column] = text

                    log(f"📤 Step {index+1}: Extracted {target_column} = '{text}'")
                    print(f"📤 Step {index+1}: Extracted {target_column} = '{text}'")

                except Exception as e:
                    print(f"⚠️ Extract skipped for {field}:\n{traceback.format_exc()}")
                    log(f"⚠️ Step {index+1}: Extract skipped for {field} — {str(e).splitlines()[0]}")
                    continue
            
            elif step["action"] == "extract_dynamic_table_flat":
                selector_key = step["selector_key"]
                selector = selector_data["selector"]  # e.g., "#datatable > tbody > tr"
                max_rows = step["params"].get("max_rows", 10)

                rows = driver.find_elements(By.CSS_SELECTOR, selector)
                print(f"📋 Found {len(rows)} rows for selector: {selector}")

                if "current_excel_row" not in context:
                    context["current_excel_row"] = {}
                current_row = context["current_excel_row"]

                extracted_count = 0
                processed = 0

                for i, row in enumerate(rows):
                    if processed >= max_rows:
                        break  # hanya ambil max_rows yang valid

                    cells = row.find_elements(By.TAG_NAME, "td")
                    cell_texts = [c.text.strip() for c in cells]

                    print(f"🧩 Row {i+1} has {len(cells)} cells")
                    for j, text in enumerate(cell_texts):
                        print(f"    [col {j+1}] {text}")

                    # Skip baris kosong total
                    if not cell_texts or all(text == "" for text in cell_texts):
                        print("⚠️ Skipping row — all cells are blank")
                        continue

                    # Skip jika tidak cukup kolom
                    if len(cells) < 5:
                        print("⚠️ Skipping row — less than 5 cells")
                        continue

                    # Skip jika baris terindikasi sebagai header
                    header_keywords = ["info code", "description", "lang desc", "details code", "seq no"]
                    if any(header.lower() in text.lower() for text in cell_texts[:5] for header in header_keywords):
                        print("⚠️ Skipping row — detected as header row")
                        continue

                    info_code = cell_texts[0]
                    suffix = info_code if info_code else f"{i+1:02d}"

                    current_row[f"InfoCode{suffix}"]     = info_code
                    current_row[f"InfoDesc{suffix}"]     = cell_texts[1]
                    current_row[f"LangDesc{suffix}"]     = cell_texts[2]
                    current_row[f"DetailsCode{suffix}"]  = cell_texts[3]
                    current_row[f"SeqNo{suffix}"]        = cell_texts[4]

                    extracted_count += 1
                    processed += 1

                context["current_excel_row"] = current_row
                print(f"✅ extract_dynamic_table_flat: extracted {extracted_count} row(s).")
                print("🔎 FINAL current_excel_row from extract_dynamic_table_flat:")
                for k, v in sorted(current_row.items()):
                    print(f"  - {k}: {v}")

            elif action == "write_hardcoded_column":
                try:
                    column_template = params.get("column", "").strip()
                    value = params.get("value")
                    n_count = context.get("last_dynamic_count", 0)

                    # Get all headers from the current Excel row
                    headers = list(context.get("current_excel_row", {}).keys())
                    print(f"[write_hardcoded_column][debug] Headers at Step {index+1}: {headers}")
                    print("[write_hardcoded_column][debug] Raw headers repr:", [repr(h) for h in headers])
                    # Normalize template for matching: collapse spaces
                    norm_template = re.sub(r"\s+", " ", column_template).strip()

                    # Escape for regex, replace {n} if present
                    if "{n}" in norm_template:
                        # Explicit pattern from user
                        pattern = re.escape(norm_template).replace(r"\{n\}", r"\d+")
                    else:
                        # Match unnumbered and numbered variants
                        pattern = rf"^{re.escape(norm_template)}(\d+)?$"

                    regex = re.compile(pattern, re.IGNORECASE)

                    # Match headers after normalizing spaces
                    matched_headers = [
                        h for h in headers
                        if regex.match(re.sub(r"\s+", " ", h).strip())
                    ]

                    if n_count:
                        matched_headers = matched_headers[:n_count]

                    # Debug log
                    print(f"[write_hardcoded_column][debug] Matching '{column_template}' → Found: {matched_headers}")

                    # Write value to matches
                    for col in matched_headers:
                        context["current_excel_row"][col] = value
                        print(f"[write_hardcoded_column] Set {col} = {value}")

                    log(f"✏️ Step {index+1}: Wrote hardcoded value '{value}' to {len(matched_headers)} columns matching '{column_template}'")

                except Exception as e:
                    log(f"❌ Step {index+1}: Failed to write hardcoded column — {e}")
                    print(f"❌ Step {index+1}: Failed to write hardcoded column — {traceback.format_exc()}")
                    continue

            else:
                log(f"❓ Step {index+1}: Unknown action '{action}'")
                print(f"❓ Step {index+1}: Unknown action '{action}'")
        
            table_index = None  # reset after each step

        except Exception as e:
            print(f"❌ Critical error in step {index+1}:\n{traceback.format_exc()}")
            log(f"❌ Step {index+1} error: {str(e).splitlines()[0]}")
            raise

    log("\n🏁 Workflow finished.")
    print("Extracted results:", results)
    log(f"✅ Final result for {context.get('Account No', '')}: {len(results.get('ATM_Table', []))} rows")
    print(f"✅ Final result for {context.get('Account No', '')}: {len(results.get('ATM_Table', []))} rows")
    return results
