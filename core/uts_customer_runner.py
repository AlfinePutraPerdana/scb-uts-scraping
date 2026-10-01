from datetime import datetime
from pathlib import Path
import re
from threading import Event

import openpyxl
from selenium.common.exceptions import NoSuchElementException, TimeoutException
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


# -----------------------------------------------------------------------------
# Workbook validation and header detection
# This section validates the Excel sheet layout before any automation runs.
# -----------------------------------------------------------------------------
FRAME_SELECTOR = "frame, iframe"
INPUT_HEADER_ALIASES = {
    2: {"CIF", "CIFNO", "CIFNUMBER", "CUSTOMERNO", "CUSTOMERNUMBER", "CUSTOMERID"},
    3: {"STATUS", "CUSTOMERSTATUS", "CHANGESTATUS", "DESIREDSTATUS", "NEWSTATUS"},
    4: {"BRANCH", "BRANCHCODE", "DESIREDBRANCH", "NEWBRANCHCODE"},
    5: {"OFFICER", "OFFICERCODE", "DESIREDOFFICER", "NEWOFFICERCODE"},
    6: {
        "CUSTOMERSEGMENT",
        "CUSTOMERSEG",
        "SEGMENT",
        "CUSTSEGMENT",
        "CUSTSEG",
        "SEGMENTCODE",
        "CUSTOMERSEGMENTCODE",
        "CUSTOMERSEGMENTTYPE",
        "GENCUSTSEGCODE",
    },
}
OUTPUT_HEADER_ALIASES = {
    7: {"RESULT", "RESULTSTATUS", "PROCESSSTATUS", "PROCESSINGSTATUS", "SCRAPESTATUS", "STATUSMSG"},
    8: {"START", "STARTTIME", "STARTDATETIME", "STARTTIMESTAMP", "PROCESSSTART"},
    9: {"END", "ENDTIME", "ENDDATETIME", "ENDTIMESTAMP", "PROCESSEND"},
}
ACCOUNT_INPUT_HEADER_ALIASES = {
    2: {"ACCOUNT", "ACCOUNTNO", "ACCOUNTNUMBER", "ACNO"},
    3: {"STATUS", "CUSTOMERSTATUS", "CHANGESTATUS", "DESIREDSTATUS", "NEWSTATUS"},
    4: {"BRANCH", "BRANCHCODE", "DESIREDBRANCH", "NEWBRANCHCODE"},
}
ACCOUNT_OUTPUT_HEADER_ALIASES = {
    5: {"RESULT", "RESULTSTATUS", "PROCESSSTATUS", "PROCESSINGSTATUS", "SCRAPESTATUS"},
    6: {"START", "STARTTIME", "STARTDATETIME", "STARTTIMESTAMP", "PROCESSSTART"},
    7: {"END", "ENDTIME", "ENDDATETIME", "ENDTIMESTAMP", "PROCESSEND"},
}
HEADER_NAMES = {
    2: "CIF_Number",
    3: "Change_Status",
    4: "Branch_Code",
    5: "Officer_Code",
    6: "Customer_Segment",
    7: "Result / Status_msg (or blank)",
    8: "Start Time (or blank)",
    9: "End Time (or blank)",
}


def normalize(value):
    return str(value or "").strip().upper()


def _normalize_header(value):
    return re.sub(r"[^A-Z0-9]", "", normalize(value))


def validate_uts_workbook(workbook_path, mode):
    """Validate the selected customer or account worksheet layout."""
    mode = str(mode).strip().lower()
    if mode not in {"maker", "checker", "account_maker", "account_checker"}:
        return False, "Choose a UTS maker or checker workflow before validating."
    is_account = mode.startswith("account_")
    input_headers = ACCOUNT_INPUT_HEADER_ALIASES if is_account else INPUT_HEADER_ALIASES
    output_headers = ACCOUNT_OUTPUT_HEADER_ALIASES if is_account else OUTPUT_HEADER_ALIASES
    header_names = (
        {
            2: "Account_Number",
            3: "Change_Status",
            4: "Branch_Code",
            5: "Result (or blank)",
            6: "Start Time (or blank)",
            7: "End Time (or blank)",
        }
        if is_account
        else HEADER_NAMES
    )

    path = Path(workbook_path)
    if not path.is_file():
        return False, "Select an existing Excel workbook."
    if path.suffix.lower() not in {".xlsx", ".xlsm", ".xls"}:
        return False, "The workbook must be an .xlsx, .xlsm, or .xls file."

    first_data_row = 5 if is_account or mode == "checker" else 4
    keep_vba = path.suffix.lower() == ".xlsm"
    try:
        workbook = openpyxl.load_workbook(
            path, read_only=True, data_only=True, keep_vba=keep_vba
        )
    except Exception as error:
        return False, f"Excel file could not be opened: {error}"

    try:
        worksheet = workbook.active
        header_row = None
        expected_header_columns = set(input_headers) | set(output_headers)
        for row_number in range(first_data_row - 1, 0, -1):
            headers = {
                column: _normalize_header(worksheet.cell(row_number, column).value)
                for column in expected_header_columns
            }
            inputs_match = all(
                headers[column] in aliases
                for column, aliases in input_headers.items()
            )
            outputs_compatible = all(
                not headers[column] or headers[column] in aliases
                for column, aliases in output_headers.items()
            )
            if inputs_match and outputs_compatible:
                header_row = row_number
                break

        if header_row is None:
            expected = ", ".join(
                f"{chr(64 + column)}: {header_names[column]}"
                for column in sorted(set(input_headers) | set(output_headers))
            )
            return False, (
                f"No compatible header row found above row {first_data_row}. "
                f"Expected {expected}."
            )

        if worksheet.cell(first_data_row, 2).value in (None, ""):
            return False, (
                f"No {'account number' if is_account else 'CIF'} was found in column B "
                f"at the first {mode} data row "
                f"({first_data_row})."
            )

        return True, (
            f"Workbook validated: sheet '{worksheet.title}', headers on row "
            f"{header_row}, {mode} data starts on row {first_data_row}."
        )
    finally:
        workbook.close()


def validate_customer_workbook(workbook_path, mode):
    mode = str(mode).strip().lower()
    if mode not in {"maker", "checker"}:
        return False, "Choose Customer Maker or Customer Checker before validating."
    return validate_uts_workbook(workbook_path, mode)


# -----------------------------------------------------------------------------
# Shared browser helpers
# These helpers keep Selenium stable while switching between frames and waiting for
# the UTS pages to load the right content before each action.
# -----------------------------------------------------------------------------
def _log(callback, message):
    if callback:
        callback(message)


def _wait_for_frame_count(driver, count, timeout):
    WebDriverWait(driver, timeout).until(
        lambda current: len(current.find_elements(By.CSS_SELECTOR, FRAME_SELECTOR)) >= count
    )


def _switch_to_details(driver, timeout=15):
    driver.switch_to.default_content()
    _wait_for_frame_count(driver, 2, timeout)
    frames = driver.find_elements(By.CSS_SELECTOR, FRAME_SELECTOR)
    driver.switch_to.frame(frames[1])
    _wait_for_frame_count(driver, 1, timeout)
    frames = driver.find_elements(By.CSS_SELECTOR, FRAME_SELECTOR)
    driver.switch_to.frame(frames[0])
    return driver


def _wait_element(driver, by, value, timeout=15):
    return WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )


def _click_element(driver, by, value, timeout=15):
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, value))
    )
    element.click()
    return element


def handle_alert(driver, timeout=1):
    """Accept a browser alert if one appears within the timeout."""
    try:
        alert = WebDriverWait(driver, timeout).until(EC.alert_is_present())
    except TimeoutException:
        return False
    alert.accept()
    return True


# -----------------------------------------------------------------------------
# Customer maker flow
# Step 1: find the customer record, read its current values, compare with Excel,
# then update the status/branch/officer fields only when the requested values differ.
# -----------------------------------------------------------------------------
def _read_customer_details(driver):
    details = {}
    for key, element_id in {
        "status": "SPAN_CUST_DETAIL_CUST_STATUS_CODE",
        "branch": "SPAN_CUST_DETAIL_GEN_BRCH_CODE",
        "officer": "SPAN_CUST_DETAIL_GEN_OFFCR_CODE",
        "customer_segment": "SPAN_CUST_DETAIL_GEN_CUST_SEG_CODE",
    }.items():
        details[key] = _wait_element(driver, By.ID, element_id).text.strip()
    return details


def _set_customer_segment(driver, value):
    segment_value = normalize(value)
    if not segment_value:
        return

    segment_select = _wait_element(driver, By.NAME, "CUST_DETAIL_GEN_CUST_SEG_CODE")
    select = Select(segment_select)
    normalized = {
        "CB": "CB",
        "CUSTOMERBANKING": "CB",
        "CUSTOMERBANKINGCB": "CB",
        "PRIORITY": "PRIORITY",
        "PRIORITYCUSTOMER": "PRIORITY",
    }
    target = normalized.get(segment_value)
    if target is None:
        raise ValueError("ERROR_segment out of scope")

    if target == "CB":
        select.select_by_index(4)
    else:
        select.select_by_index(20)


def _search_customer_maker(driver, cif):
    _switch_to_details(driver)
    search_input = _wait_element(driver, By.ID, "SEARCH_TEXT")
    search_input.clear()
    search_input.send_keys(str(cif).strip())
    _click_element(driver, By.CSS_SELECTOR, ".td_search_btn .btn.toggleButton")

    _switch_to_details(driver)
    result = _wait_element(driver, By.ID, "SPAN_CUST_DETAIL_LIST_CUST_NO_0")
    result.click()
    _switch_to_details(driver)


def _run_maker_row(driver, values):
    cif, requested_status, requested_branch, requested_officer, requested_segment = values
    _search_customer_maker(driver, cif)
    current = _read_customer_details(driver)
    requested = {
        "status": requested_status,
        "branch": requested_branch,
        "officer": requested_officer,
        "customer_segment": requested_segment,
    }
    updates = {
        key: value
        for key, value in requested.items()
        if normalize(value) and normalize(value) != normalize(current[key])
    }
    if not updates:
        return "No changes"

    edit_link = _wait_element(driver, By.CSS_SELECTOR, ".td_serach_btn a")
    edit_link.click()
    _switch_to_details(driver)

    if "status" in updates:
        status_select = _wait_element(driver, By.NAME, "CUST_DETAIL_CUST_STATUS_CODE")
        Select(status_select).select_by_visible_text(str(updates["status"]).strip())
    if "branch" in updates:
        branch_input = _wait_element(driver, By.ID, "CUST_DETAIL_GEN_BRCH_CODE")
        branch_input.clear()
        branch_input.send_keys(str(updates["branch"]).strip())
    if "officer" in updates:
        officer_input = _wait_element(driver, By.ID, "CUST_DETAIL_GEN_OFFCR_CODE")
        officer_input.clear()
        officer_input.send_keys(str(updates["officer"]).strip())
    if "customer_segment" in updates:
        _set_customer_segment(driver, updates["customer_segment"])

    _click_element(driver, By.CSS_SELECTOR, ".td_serach_btn a")
    return "Done"


# -----------------------------------------------------------------------------
# Account maker flow
# This flow handles account record changes and status transitions before saving.
# -----------------------------------------------------------------------------
def _account_action_link(driver, index):
    links = driver.find_elements(By.CSS_SELECTOR, ".de_but a")
    if len(links) <= index:
        raise NoSuchElementException(f"Account action link {index} was not found")
    return links[index]


def _search_account_maker(driver, account_number):
    _switch_to_details(driver)
    search_input = _wait_element(driver, By.ID, "SEARCH_TEXT")
    search_input.clear()
    search_input.send_keys(str(account_number).strip())
    _click_element(driver, By.CSS_SELECTOR, ".td_search_btn .btn.toggleButton")

    _switch_to_details(driver)
    _wait_element(driver, By.ID, "SPAN_ACCOUNT_LIST_ACA_AC_NO_0").click()
    _switch_to_details(driver)


def _read_account_details(driver):
    return {
        "status": _wait_element(
            driver, By.ID, "SPAN_ACCOUNT_DETAIL_CUST_STATUS_CODE"
        ).text.strip(),
        "branch": _wait_element(
            driver, By.ID, "SPAN_ACCOUNT_DETAIL_GEN_BRCH_CODE"
        ).text.strip(),
    }


def _submit_account_status(driver, status):
    modal = _wait_element(driver, By.NAME, "myframe__1")
    driver.switch_to.frame(modal)
    status_select = _wait_element(driver, By.NAME, "new_cust_status_code")
    Select(status_select).select_by_visible_text(str(status).strip())
    ok_button = next(
        (
            button
            for button in driver.find_elements(By.CSS_SELECTOR, ".btn.toggleButton")
            if normalize(button.text) == "OK"
        ),
        None,
    )
    if ok_button is None:
        raise NoSuchElementException("Account status confirmation button was not found")
    ok_button.click()
    driver.switch_to.default_content()
    _switch_to_details(driver)
    _wait_element(driver, By.ID, "SPAN_ACCOUNT_DETAIL_CUST_STATUS_CODE")


def _change_account_status(driver, current_status, requested_status):
    if normalize(current_status) == "CLOSED":
        _account_action_link(driver, 2).click()
        _submit_account_status(driver, "DMI")
        _account_action_link(driver, 3).click()
    else:
        _account_action_link(driver, 2).click()
    _submit_account_status(driver, requested_status)


def _run_account_maker_row(driver, values):
    account_number, requested_status, requested_branch = values[:3]
    _search_account_maker(driver, account_number)
    current = _read_account_details(driver)
    updates = {
        key: value
        for key, value in {
            "status": requested_status,
            "branch": requested_branch,
        }.items()
        if normalize(value) and normalize(value) != normalize(current[key])
    }
    if not updates:
        return "No changes"

    if "status" in updates:
        _change_account_status(driver, current["status"], updates["status"])
    if "branch" in updates:
        _account_action_link(driver, 0).click()
        _switch_to_details(driver)
        branch_input = _wait_element(driver, By.ID, "ACCOUNT_DETAIL_GEN_BRCH_CODE")
        branch_input.clear()
        branch_input.send_keys(str(updates["branch"]).strip())
        _account_action_link(driver, 0).click()
    return "Done"


# -----------------------------------------------------------------------------
# Customer checker flow
# This section opens a record in checker mode, reads the current values from the
# detail screen, and decides whether the record matches the workbook expectations.
# -----------------------------------------------------------------------------
def _open_checker_record(driver, cif):
    _switch_to_details(driver)
    search_buttons = driver.find_elements(By.CSS_SELECTOR, ".btn.toggleButton")
    search_button = next(
        (button for button in search_buttons if normalize(button.text) == "SEARCH"),
        None,
    )
    if search_button is None:
        raise NoSuchElementException("Checker Search button was not found")
    search_button.click()

    _switch_to_details(driver)
    modal = _wait_element(driver, By.NAME, "myframe__1")
    driver.switch_to.frame(modal)
    cif_input = _wait_element(driver, By.ID, "SEARCH_CUST_NO")
    cif_input.clear()
    cif_input.send_keys(str(cif).strip())
    _click_element(driver, By.CSS_SELECTOR, ".btn.toggleButton")
    driver.switch_to.parent_frame()

    _switch_to_details(driver)
    try:
        _wait_element(driver, By.CSS_SELECTOR, "form[name='fmDetails'] table", timeout=8)
    except TimeoutException:
        return False
    rows = driver.find_elements(By.CSS_SELECTOR, "form[name='fmDetails'] table tr")
    if len(rows) <= 2:
        _click_refresh_if_present(driver)
        return False
    links = rows[2].find_elements(By.CSS_SELECTOR, "td:nth-child(2) a")
    if not links:
        _click_refresh_if_present(driver)
        return False
    links[0].click()
    return True


def _click_refresh_if_present(driver):
    for button in driver.find_elements(By.CSS_SELECTOR, ".btn.toggleButton"):
        if normalize(button.text) == "REFRESH":
            button.click()
            return


def _read_checker_record(driver, expected_values):
    _switch_to_details(driver)
    modal = _wait_element(driver, By.NAME, "myframe__1")
    driver.switch_to.frame(modal)
    dummy = _wait_element(driver, By.ID, "dummy")
    current = {
        "status": dummy.find_element(By.ID, "SPAN_CUST_DETAIL_CUST_STATUS_CODE").text.strip(),
        "branch": dummy.find_element(By.ID, "SPAN_CUST_DETAIL_GEN_BRCH_CODE").text.strip(),
        "officer": dummy.find_element(By.ID, "SPAN_CUST_DETAIL_GEN_OFFCR_CODE").text.strip(),
        "customer_segment": dummy.find_element(
            By.ID, "SPAN_CUST_DETAIL_GEN_CUST_SEG_CODE"
        ).text.strip(),
    }
    mismatches = []
    fields = {
        "change_status": "status",
        "branch_code": "branch",
        "officer_code": "officer",
        "customer_segment": "customer_segment",
    }
    for column_name, field in fields.items():
        expected = expected_values.get(column_name, "")
        if normalize(expected) and normalize(expected) != normalize(current[field]):
            mismatches.append(column_name)

    close_buttons = driver.find_elements(By.CSS_SELECTOR, ".btn.toggleButton")
    for button in close_buttons:
        if normalize(button.text) == "CLOSE":
            button.click()
            break
    driver.switch_to.default_content()

    if mismatches:
        return ", ".join(mismatches)

    # The source macro initializes hasOfficerCode=True, making its decision
    # unconditional. Do not automate Approve/Undo until that rule is confirmed.
    return "No mismatches - manual approval required"


# -----------------------------------------------------------------------------
# Account checker flow
# Similar to customer checker logic, but for account-level records and account
# approval/undo recommendation checks.
# -----------------------------------------------------------------------------
def _open_account_checker_record(driver, account_number):
    _switch_to_details(driver)
    search_button = next(
        (
            button
            for button in driver.find_elements(By.CSS_SELECTOR, ".btn.toggleButton")
            if normalize(button.text) == "SEARCH"
        ),
        None,
    )
    if search_button is None:
        raise NoSuchElementException("Account checker Search button was not found")
    search_button.click()

    _switch_to_details(driver)
    modal = _wait_element(driver, By.NAME, "myframe__1")
    driver.switch_to.frame(modal)
    account_input = _wait_element(driver, By.ID, "SEARCH_ACA_AC_NO")
    account_input.clear()
    account_input.send_keys(str(account_number).strip())
    ok_button = next(
        (
            button
            for button in driver.find_elements(By.CSS_SELECTOR, ".btn.toggleButton")
            if normalize(button.text) == "OK"
        ),
        None,
    )
    if ok_button is None:
        raise NoSuchElementException("Account checker confirmation button was not found")
    ok_button.click()
    driver.switch_to.default_content()

    _switch_to_details(driver)
    try:
        _wait_element(driver, By.CSS_SELECTOR, "form[name='fmDetails'] table", timeout=8)
    except TimeoutException:
        _click_refresh_if_present(driver)
        return False
    rows = driver.find_elements(By.CSS_SELECTOR, "form[name='fmDetails'] table tr")
    if len(rows) <= 2:
        _click_refresh_if_present(driver)
        return False
    links = rows[2].find_elements(By.CSS_SELECTOR, "td:nth-child(2) a")
    if not links:
        _click_refresh_if_present(driver)
        return False
    links[0].click()
    _switch_to_details(driver)
    return True


def _is_red_css_color(color):
    normalized_color = normalize(color).replace(" ", "")
    return normalized_color in {"RED", "#FF0000", "RGB(255,0,0)", "RGBA(255,0,0,1)"}


def _read_account_checker_record(driver, expected_values):
    _switch_to_details(driver)
    modal = _wait_element(driver, By.NAME, "myframe__1")
    driver.switch_to.frame(modal)
    dummy = _wait_element(driver, By.ID, "dummy")
    current = {
        "status": dummy.find_element(
            By.ID, "SPAN_ACCOUNT_DETAIL_CUST_STATUS_CODE"
        ).text.strip(),
        "branch": dummy.find_element(
            By.ID, "SPAN_ACCOUNT_DETAIL_GEN_BRCH_CODE"
        ).text.strip(),
    }
    mismatches = []
    for column_name, field in (("Change_Status", "status"), ("Branch_Code", "branch")):
        expected = expected_values.get(field, "")
        if normalize(expected) and normalize(expected) != normalize(current[field]):
            mismatches.append(column_name)

    approve_data = False
    for label in driver.find_elements(By.CSS_SELECTOR, ".tablabel.th"):
        fonts = label.find_elements(By.TAG_NAME, "font")
        if not fonts:
            continue
        label_text = normalize(fonts[0].text)
        if label_text in {"ACCOUNT STATUS", "BRANCH"} and _is_red_css_color(
            fonts[0].value_of_css_property("color")
        ):
            approve_data = True

    close_button = _wait_element(driver, By.CSS_SELECTOR, ".close")
    close_button.click()
    driver.switch_to.default_content()
    if mismatches:
        return ", ".join(mismatches)
    recommendation = "Approve" if approve_data else "Undo"
    return f"No mismatches - {recommendation} manually"


# -----------------------------------------------------------------------------
# Row processors
# These are the actual per-row automation actions called by the top-level workflow.
# Each function receives the Excel row values and executes a single maker/checker step.
# -----------------------------------------------------------------------------
def _run_checker_row(driver, values):
    cif = values[0]
    expected_values = {
        "change_status": values[1],
        "branch_code": values[2],
        "officer_code": values[3],
        "customer_segment": values[4],
    }
    if not _open_checker_record(driver, cif):
        return "CIF Not Found"
    return _read_checker_record(driver, expected_values)


def _run_account_checker_row(driver, values):
    expected_values = {"status": values[1], "branch": values[2]}
    if not _open_account_checker_record(driver, values[0]):
        return "Account Not Found"
    return _read_account_checker_record(driver, expected_values)


# -----------------------------------------------------------------------------
# Workbook runner
# This is the main orchestration loop: it reads each Excel row, marks start time,
# executes the appropriate maker/checker row processor, writes the result, and saves.
# -----------------------------------------------------------------------------
def run_uts_workflow(driver, workbook_path, mode, log_callback=None, stop_event=None):
    """Run a customer or account maker/checker workflow from its Excel rows."""
    mode = mode.lower().strip()
    if mode not in {"maker", "checker", "account_maker", "account_checker"}:
        raise ValueError("mode must be maker, checker, account_maker, or account_checker")

    path = Path(workbook_path)
    keep_vba = path.suffix.lower() == ".xlsm"
    workbook = openpyxl.load_workbook(path, keep_vba=keep_vba)
    worksheet = workbook.active
    is_account = mode.startswith("account_")
    is_maker = mode in {"maker", "account_maker"}
    first_row = 5 if is_account or not is_maker else 4
    result_column, start_column, end_column = (5, 6, 7) if is_account else (7, 8, 9)
    processor = {
        "maker": _run_maker_row,
        "checker": _run_checker_row,
        "account_maker": _run_account_maker_row,
        "account_checker": _run_account_checker_row,
    }[mode]
    stop_event = stop_event or Event()

    try:
        for row_number in range(first_row, worksheet.max_row + 1):
            cif = worksheet.cell(row_number, 2).value
            if cif in (None, ""):
                break
            if worksheet.cell(row_number, result_column).value not in (None, ""):
                _log(log_callback, f"Row {row_number}: skipped (already has a result).")
                continue
            if stop_event.is_set():
                _log(log_callback, "Stop requested; leaving remaining rows untouched.")
                break

            worksheet.cell(row_number, start_column, datetime.now())
            try:
                columns_to_read = range(2, 7) if not is_account else range(2, 6)
                values = [worksheet.cell(row_number, column).value for column in columns_to_read]
                result = processor(driver, values)
                worksheet.cell(row_number, result_column, result)
                _log(log_callback, f"Row {row_number}: {result}")
            except Exception as error:
                worksheet.cell(row_number, result_column, f"Error: {error}")
                _log(log_callback, f"Row {row_number}: Error: {error}")
            finally:
                worksheet.cell(row_number, end_column, datetime.now())
                workbook.save(path)
    finally:
        workbook.close()


def run_customer_workflow(driver, workbook_path, mode, log_callback=None, stop_event=None):
    mode = str(mode).strip().lower()
    if mode not in {"maker", "checker"}:
        raise ValueError("mode must be 'maker' or 'checker'")
    return run_uts_workflow(
        driver, workbook_path, mode, log_callback=log_callback, stop_event=stop_event
    )
