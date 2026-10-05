from datetime import datetime
from contextvars import ContextVar
from pathlib import Path
import re
from threading import Event
import traceback

import openpyxl
from selenium.common.exceptions import (
    NoSuchElementException,
    NoSuchWindowException,
    TimeoutException,
    UnexpectedAlertPresentException,
)
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select, WebDriverWait


# -----------------------------------------------------------------------------
# Workbook validation and header detection
# This section validates the Excel sheet layout before any automation runs.
# -----------------------------------------------------------------------------
FRAME_SELECTOR = "frame, iframe"
UTS_WINDOW_TITLE = "Unit Trust System"
_WORKFLOW_LOG_CALLBACK = ContextVar("uts_workflow_log_callback", default=None)
INPUT_HEADER_ALIASES = {
    1: {"RELID"},
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
    7: {"STATUSMSG"},
    8: {"STARTDATE"},
    9: {"ENDDATE"},
}
ACCOUNT_INPUT_HEADER_ALIASES = {
    1: {"RELID"},
    2: {"ACCOUNT", "ACCOUNTNO", "ACCOUNTNUMBER", "PRODUCTACCOUNT", "ACNO"},
    3: {"STATUS", "CUSTOMERSTATUS", "CHANGESTATUS", "DESIREDSTATUS", "NEWSTATUS"},
    4: {"BRANCH", "BRANCHCODE", "DESIREDBRANCH", "NEWBRANCHCODE"},
}
ACCOUNT_OUTPUT_HEADER_ALIASES = {
    5: {"STATUSMSG"},
    6: {"STARTDATE"},
    7: {"ENDDATE"},
}
HEADER_NAMES = {
    1: "Rel_ID",
    2: "CIF_Number",
    3: "Change_Status",
    4: "Branch_Code",
    5: "Officer_Code",
    6: "Customer_Segment",
    7: "Status_msg",
    8: "Start_Date",
    9: "End_Date",
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
            1: "Rel_ID",
            2: "Product_Account",
            3: "Change_Status",
            4: "Branch_Code",
            5: "Status_msg",
            6: "Start_Date",
            7: "End_Date",
        }
        if is_account
        else HEADER_NAMES
    )

    path = Path(workbook_path)
    if not path.is_file():
        return False, "Select an existing Excel workbook."
    if path.suffix.lower() not in {".xlsx", ".xlsm"}:
        return False, "The workbook must be an .xlsx or .xlsm file."

    header_row = 4 if is_account else 5
    first_data_row = header_row + 1
    keep_vba = path.suffix.lower() == ".xlsm"
    try:
        workbook = openpyxl.load_workbook(
            path, read_only=True, data_only=True, keep_vba=keep_vba
        )
    except Exception as error:
        return False, f"Excel file could not be opened: {error}"

    try:
        worksheet = workbook.active
        expected_header_columns = set(input_headers) | set(output_headers)
        headers = {
            column: _normalize_header(worksheet.cell(header_row, column).value)
            for column in expected_header_columns
        }
        inputs_match = all(
            headers[column] in aliases
            for column, aliases in input_headers.items()
        )
        outputs_match = all(
            headers[column] in aliases
            for column, aliases in output_headers.items()
        )
        if not inputs_match or not outputs_match:
            expected = ", ".join(
                f"{chr(64 + column)}: {header_names[column]}"
                for column in sorted(set(input_headers) | set(output_headers))
            )
            return False, (
                f"No compatible header row found on row {header_row}. "
                f"Expected {expected}."
            )

        has_identifier = worksheet.cell(first_data_row, 2).value not in (None, "")
        if not is_account and mode == "checker":
            has_identifier = has_identifier or worksheet.cell(first_data_row, 1).value not in (None, "")
        if not has_identifier:
            identifier_name = "account number" if is_account else (
                "CIF or Rel_ID" if mode == "checker" else "CIF"
            )
            return False, (
                f"No {identifier_name} was found at the first {mode} data row "
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
    callback = callback or _WORKFLOW_LOG_CALLBACK.get()
    if callback:
        callback(message)


def _switch_to_uts_window(driver, timeout=30):
    """Select the UTS popup opened from the login window by its browser title."""
    observed_titles = set()
    reported_titles = set()
    _log(None, f"Looking for a browser window titled '{UTS_WINDOW_TITLE}'.")

    def find_uts_window(current_driver):
        for handle in current_driver.window_handles:
            try:
                current_driver.switch_to.window(handle)
                title = current_driver.title
            except NoSuchWindowException:
                continue
            observed_titles.add(title or "<untitled>")
            if title not in reported_titles:
                _log(None, f"Found browser window title {title or '<untitled>'!r}.")
                reported_titles.add(title)
            if UTS_WINDOW_TITLE.casefold() in (title or "").casefold():
                _log(None, f"Selected UTS window: handle={handle}, title={title!r}.")
                return handle
        return False

    try:
        return WebDriverWait(driver, timeout, poll_frequency=0.5).until(
            find_uts_window
        )
    except TimeoutException as error:
        titles = ", ".join(sorted(observed_titles)) or "<no browser windows>"
        raise TimeoutException(
            f"Could not find a browser window titled '{UTS_WINDOW_TITLE}'. "
            "After signing in, open UTS from its menu in the WebDriver browser "
            f"and try again. Browser titles found: {titles}"
        ) from error


def _wait_for_frame_count(driver, count, timeout):
    _log(None, f"Waiting for at least {count} frame(s) (timeout={timeout}s).")
    WebDriverWait(driver, timeout).until(
        lambda current: len(current.find_elements(By.CSS_SELECTOR, FRAME_SELECTOR)) >= count
    )
    _log(None, f"Found at least {count} frame(s).")


def _switch_to_details(driver, timeout=15):
    _log(None, "Switching to top-level page content.")
    driver.switch_to.default_content()
    _wait_for_frame_count(driver, 2, timeout)
    frames = driver.find_elements(By.CSS_SELECTOR, FRAME_SELECTOR)
    _log(None, "Switching to the UTS content frame (frame index 1).")
    driver.switch_to.frame(frames[1])
    _wait_for_frame_count(driver, 1, timeout)
    frames = driver.find_elements(By.CSS_SELECTOR, FRAME_SELECTOR)
    _log(None, "Switching to the UTS details frame (nested frame index 0).")
    driver.switch_to.frame(frames[0])
    return driver


def _wait_element(driver, by, value, timeout=15):
    _log(None, f"Waiting for element {by}={value!r} (timeout={timeout}s).")
    element = WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((by, value))
    )
    _log(None, f"Found element {by}={value!r}.")
    return element


def _click_element(driver, by, value, timeout=15):
    _log(None, f"Waiting for clickable element {by}={value!r}.")
    element = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((by, value))
    )
    element.click()
    _log(None, f"Clicked element {by}={value!r}.")
    return element


def _click_parent_link(driver, parent_selector, index, action, timeout=15):
    _log(
        None,
        f"Waiting for {action} link {index} inside parent {parent_selector!r}.",
    )

    def find_link(current_driver):
        parent = current_driver.find_element(By.CSS_SELECTOR, parent_selector)
        links = parent.find_elements(By.TAG_NAME, "a")
        if len(links) <= index:
            return False
        link = links[index]
        if link.is_displayed() and link.is_enabled():
            return link
        return False

    link = WebDriverWait(driver, timeout).until(find_link)
    _log(None, f"{action} link text: {link.text.strip()!r}.")
    link.click()
    _log(None, f"Clicked {action} link in parent {parent_selector!r}.")
    return link


def _click_button_by_text(driver, expected_text, selector=".btn.toggleButton", timeout=15):
    expected = normalize(expected_text)
    _log(None, f"Looking for a unique button labeled {expected_text!r}.")

    def find_buttons(current_driver):
        buttons = [
            button
            for button in current_driver.find_elements(By.CSS_SELECTOR, selector)
            if normalize(button.text) == expected
            and button.is_displayed()
            and button.is_enabled()
        ]
        return buttons or False

    buttons = WebDriverWait(driver, timeout).until(find_buttons)
    if len(buttons) != 1:
        raise NoSuchElementException(
            f"Expected one button labeled {expected_text!r}, found {len(buttons)}."
        )
    button = buttons[0]
    _log(None, f"Clicked button labeled {button.text.strip()!r}.")
    button.click()
    return button


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
    _log(None, "Reading current customer status, branch, officer, and segment.")
    details = {}
    for key, element_id in {
        "status": "SPAN_CUST_DETAIL_CUST_STATUS_CODE",
        "branch": "SPAN_CUST_DETAIL_GEN_BRCH_CODE",
        "officer": "SPAN_CUST_DETAIL_GEN_OFFCR_CODE",
        "customer_segment": "SPAN_CUST_DETAIL_GEN_CUST_SEG_CODE",
    }.items():
        details[key] = _wait_element(driver, By.ID, element_id).text.strip()
        _log(None, f"Current customer {key}: {details[key]!r}.")
    return details


def _set_customer_segment(driver, value):
    target = _customer_segment_target(value)
    if target is None:
        return
    segment_select = _wait_element(driver, By.NAME, "CUST_DETAIL_GEN_CUST_SEG_CODE")
    select = Select(segment_select)
    select.select_by_index(4 if target == "CB" else 20)
    _log(None, f"Selected customer segment {target!r}.")


def _customer_segment_target(value):
    segment_value = normalize(value)
    if not segment_value:
        return None
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
    return target


def _search_customer_maker(driver, cif):
    _log(None, f"Customer Maker: searching CIF {cif!r}.")
    _switch_to_details(driver)
    search_input = _wait_element(driver, By.ID, "SEARCH_TEXT")
    search_input.clear()
    search_input.send_keys(str(cif).strip())
    _log(None, "Entered CIF in SEARCH_TEXT.")
    _click_element(
        driver,
        By.CSS_SELECTOR,
        "#container > section > div > div > form > div:nth-child(4) > table > "
        "tbody > tr:nth-child(1) > td.td_search_btn > a:nth-child(1)",
    )

    _switch_to_details(driver)
    result = _wait_element(driver, By.ID, "SPAN_CUST_DETAIL_LIST_CUST_NO_0")
    result.click()
    _log(None, "Opened the first matching customer record.")
    _switch_to_details(driver)


def _run_maker_row(driver, values):
    cif, requested_status, requested_branch, requested_officer, requested_segment = values
    _log(None, f"Customer Maker requested values: {values!r}.")
    if normalize(requested_segment):
        _customer_segment_target(requested_segment)

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
    _log(None, f"Customer Maker fields requiring updates: {updates!r}.")
    if not updates:
        _log(None, "Customer values already match; skipping edit.")
        return "No changes"

    _log(None, "Opening Customer Maker edit form.")
    _click_parent_link(driver, ".td_serach_btn", 0, "Customer Edit")
    _switch_to_details(driver)

    editable_field_ids = {
        "status": "CUST_DETAIL_CUST_STATUS_CODE",
        "branch": "CUST_DETAIL_GEN_BRCH_CODE",
        "officer": "CUST_DETAIL_GEN_OFFCR_CODE",
        "customer_segment": "CUST_DETAIL_GEN_CUST_SEG_CODE",
    }
    for key in updates:
        _wait_element(driver, By.ID, editable_field_ids[key])
        _log(None, f"Customer edit field is ready: {editable_field_ids[key]}.")

    if "status" in updates:
        status_select = _wait_element(
            driver, By.ID, "CUST_DETAIL_CUST_STATUS_CODE"
        )
        Select(status_select).select_by_visible_text(str(updates["status"]).strip())
        _log(None, f"Set customer status to {updates['status']!r}.")
    if "branch" in updates:
        branch_input = _wait_element(driver, By.ID, "CUST_DETAIL_GEN_BRCH_CODE")
        branch_input.clear()
        branch_input.send_keys(str(updates["branch"]).strip())
        _log(None, f"Entered customer branch {updates['branch']!r}.")
    if "officer" in updates:
        officer_input = _wait_element(driver, By.ID, "CUST_DETAIL_GEN_OFFCR_CODE")
        officer_input.clear()
        officer_input.send_keys(str(updates["officer"]).strip())
        _log(None, f"Entered customer officer {updates['officer']!r}.")
    if "customer_segment" in updates:
        _set_customer_segment(driver, updates["customer_segment"])

    _log(None, "Saving Customer Maker edits.")
    _click_parent_link(driver, ".td_serach_btn", 0, "Customer Save")
    _switch_to_details(driver)
    display_field_ids = {
        "status": "SPAN_CUST_DETAIL_CUST_STATUS_CODE",
        "branch": "SPAN_CUST_DETAIL_GEN_BRCH_CODE",
        "officer": "SPAN_CUST_DETAIL_GEN_OFFCR_CODE",
        "customer_segment": "SPAN_CUST_DETAIL_GEN_CUST_SEG_CODE",
    }
    for key, expected in updates.items():
        actual = _wait_element(driver, By.ID, display_field_ids[key]).text.strip()
        if normalize(actual) != normalize(expected):
            raise RuntimeError(
                f"Customer save verification failed for {key}: "
                f"expected {expected!r}, found {actual!r}."
            )
        _log(None, f"Verified saved customer {key}: {actual!r}.")
    _log(None, "Customer Maker save action completed and verified.")
    return "Done"


# -----------------------------------------------------------------------------
# Account maker flow
# This flow handles account record changes and status transitions before saving.
# -----------------------------------------------------------------------------
def _account_action_link(driver, index):
    parent = driver.find_element(By.CSS_SELECTOR, ".de_but")
    links = parent.find_elements(By.TAG_NAME, "a")
    if len(links) <= index:
        raise NoSuchElementException(f"Account action link {index} was not found")
    _log(
        None,
        f"Found account action link index {index} of {len(links)}; "
        f"text={links[index].text.strip()!r}.",
    )
    return links[index]


def _search_account_maker(driver, account_number):
    _log(None, f"Account Maker: searching account {account_number!r}.")
    _switch_to_details(driver)
    search_input = _wait_element(driver, By.ID, "SEARCH_TEXT")
    search_input.clear()
    search_input.send_keys(str(account_number).strip())
    _log(None, "Entered account number in SEARCH_TEXT.")
    _click_element(
        driver,
        By.CSS_SELECTOR,
        "#container > section > div > div > form > div:nth-child(2) > table > "
        "tbody > tr:nth-child(1) > td.td_search_btn > a:nth-child(1)",
    )

    _switch_to_details(driver)
    _wait_element(driver, By.ID, "SPAN_ACCOUNT_LIST_ACA_AC_NO_0").click()
    _log(None, "Opened the first matching account record.")
    _switch_to_details(driver)


def _read_account_details(driver):
    _log(None, "Reading current account status and branch.")
    details = {
        "status": _wait_element(
            driver, By.ID, "SPAN_ACCOUNT_DETAIL_CUST_STATUS_CODE"
        ).text.strip(),
        "branch": _wait_element(
            driver, By.ID, "SPAN_ACCOUNT_DETAIL_GEN_BRCH_CODE"
        ).text.strip(),
    }
    _log(None, f"Current account details: {details!r}.")
    return details


def _submit_account_status(driver, status):
    _log(None, f"Opening account status dialog; requested status={status!r}.")
    modal = _wait_element(driver, By.NAME, "myframe__1")
    driver.switch_to.frame(modal)
    status_select = _wait_element(driver, By.CSS_SELECTOR, "#fmStatus select")
    Select(status_select).select_by_visible_text(str(status).strip())
    _log(None, f"Selected account status {status!r}; submitting dialog.")
    try:
        _click_element(
            driver, By.CSS_SELECTOR, "#fmStatus .ppw_foot_btn a:nth-child(1)"
        )
    except UnexpectedAlertPresentException:
        alert = driver.switch_to.alert
    else:
        try:
            alert = WebDriverWait(driver, 1).until(EC.alert_is_present())
        except TimeoutException:
            alert = None
    if alert is not None:
        alert_message = alert.text
        _log(None, f"Account status dialog raised alert: {alert_message!r}.")
        alert.accept()
        _log(None, "Accepted status alert; cancelling the status change.")
        _click_element(
            driver, By.CSS_SELECTOR, "#fmStatus .ppw_foot_btn a:nth-child(2)"
        )
        raise RuntimeError(f"Account status change cancelled: {alert_message}")
    driver.switch_to.default_content()
    _switch_to_details(driver)
    _wait_element(driver, By.ID, "SPAN_ACCOUNT_DETAIL_CUST_STATUS_CODE")
    _log(None, f"Account status change to {status!r} submitted.")


def _change_account_status(driver, current_status, requested_status):
    _log(
        None,
        f"Changing account status from {current_status!r} to {requested_status!r}.",
    )
    if normalize(current_status) == "CLOSED":
        _log(None, "Account is closed; switching through DMI before reopening.")
        _account_action_link(driver, 2).click()
        _submit_account_status(driver, "DMI")
        _account_action_link(driver, 3).click()
    else:
        _account_action_link(driver, 2).click()
    _submit_account_status(driver, requested_status)


def _run_account_maker_row(driver, values):
    account_number, requested_status, requested_branch = values[:3]
    _log(None, f"Account Maker requested values: {values!r}.")
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
    _log(None, f"Account Maker fields requiring updates: {updates!r}.")
    if not updates:
        _log(None, "Account values already match; skipping edit.")
        return "No changes"

    if "status" in updates:
        _change_account_status(driver, current["status"], updates["status"])
    if "branch" in updates:
        _log(None, "Opening Account Maker branch edit form.")
        _account_action_link(driver, 0).click()
        _switch_to_details(driver)
        branch_input = _wait_element(driver, By.ID, "ACCOUNT_DETAIL_GEN_BRCH_CODE")
        _log(None, "Account branch edit field is ready.")
        branch_input.clear()
        branch_input.send_keys(str(updates["branch"]).strip())
        _log(None, f"Entered account branch {updates['branch']!r}; saving.")
        _account_action_link(driver, 0).click()
        _switch_to_details(driver)
        actual_branch = _wait_element(
            driver, By.ID, "SPAN_ACCOUNT_DETAIL_GEN_BRCH_CODE"
        ).text.strip()
        if normalize(actual_branch) != normalize(updates["branch"]):
            raise RuntimeError(
                "Account save verification failed for branch: "
                f"expected {updates['branch']!r}, found {actual_branch!r}."
            )
        _log(None, f"Verified saved account branch: {actual_branch!r}.")
        _log(None, "Account Maker branch save action completed.")
    return "Done"


# -----------------------------------------------------------------------------
# Customer checker flow
# This section opens a record in checker mode, reads the current values from the
# detail screen, and decides whether the record matches the workbook expectations.
# -----------------------------------------------------------------------------
def _open_checker_record(driver, cif, rel_id=None):
    _log(None, f"Customer Checker: searching by CIF={cif!r}, Rel_ID={rel_id!r}.")
    _switch_to_details(driver)
    _click_button_by_text(driver, "Search")

    _switch_to_details(driver)
    modal = _wait_element(driver, By.NAME, "myframe__1")
    driver.switch_to.frame(modal)
    if normalize(cif):
        search_input = _wait_element(driver, By.ID, "SEARCH_CUST_NO")
        search_value = cif
        _log(None, "Customer Checker selected CIF search (CIF takes priority).")
    else:
        search_input = _wait_element(driver, By.ID, "SEARCH_CUST_EBBS_CAT_NO")
        search_value = rel_id
        _log(None, "Customer Checker selected Rel_ID search.")
    search_input.clear()
    search_input.send_keys(str(search_value).strip())
    _log(None, f"Entered checker search value {search_value!r}.")
    _click_button_by_text(driver, "OK", selector="nav > a")
    _log(None, "Submitted Customer Checker search.")
    driver.switch_to.parent_frame()

    _switch_to_details(driver)
    try:
        _wait_element(driver, By.CSS_SELECTOR, "form[name='fmDetails'] table", timeout=8)
    except TimeoutException:
        _log(None, "Customer Checker result table did not appear.")
        return False
    rows = driver.find_elements(By.CSS_SELECTOR, "form[name='fmDetails'] table tr")
    if len(rows) <= 2:
        _log(None, "Customer Checker search returned no record row.")
        _click_refresh_if_present(driver)
        return False
    links = driver.find_elements(
        By.CSS_SELECTOR,
        "#container > section > div > form > table > tbody > "
        "tr.gridcolumneven > td:nth-child(2) > a",
    )
    if not links:
        _log(None, "Customer Checker result Function ID link not found.")
        _click_refresh_if_present(driver)
        return False
    links[0].click()
    _log(None, "Opened Customer Checker Function ID record.")
    return True


def _click_refresh_if_present(driver):
    for button in driver.find_elements(By.CSS_SELECTOR, ".btn.toggleButton"):
        if normalize(button.text) == "REFRESH":
            button.click()
            _log(None, "Clicked Refresh after an empty checker search.")
            return


def _read_checker_record(driver, expected_values):
    _log(None, f"Reading Customer Checker record; expected values={expected_values!r}.")
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
        "Change_Status": ("change_status", "status"),
        "Branch code": ("branch_code", "branch"),
        "Officer code": ("officer_code", "officer"),
        "Customer Segment": ("customer_segment", "customer_segment"),
    }
    for label, (column_name, field) in fields.items():
        expected = expected_values.get(column_name, "")
        if normalize(expected) and normalize(expected) != normalize(current[field]):
            _log(
                None,
                f"Customer Checker mismatch for {label}: expected={expected!r}, "
                f"actual={current[field]!r}.",
            )
            mismatches.append(
                f"{label} {expected} found {current[field]} instead"
            )
        elif normalize(expected):
            _log(None, f"Customer Checker {label} matches ({current[field]!r}).")
        else:
            _log(None, f"Customer Checker {label} is blank in workbook; ignored.")

    _click_button_by_text(driver, "Close")
    driver.switch_to.default_content()

    if mismatches:
        _log(None, f"Customer Checker mismatches found: {mismatches!r}.")
        return "MISMATCH : " + "; ".join(mismatches)
    _log(None, "Customer Checker values all match.")
    return ""


def _approve_customer_checker_record(driver):
    _log(None, "Customer Checker matched; starting close-and-approve steps.")
    _switch_to_details(driver)
    _click_parent_link(
        driver,
        "#container > section > div.h_tab > div.hom_h > form > div > div",
        0,
        "Customer Close",
    )
    _wait_element(driver, By.ID, "MarkDel").click()
    _log(None, "Checked MarkDel checklist.")
    _click_button_by_text(driver, "Approve")


# -----------------------------------------------------------------------------
# Account checker flow
# Similar to customer checker logic, but for account-level records and account
# approval/undo recommendation checks.
# -----------------------------------------------------------------------------
def _open_account_checker_record(driver, account_number):
    _log(None, f"Account Checker: searching account {account_number!r}.")
    _switch_to_details(driver)
    _click_button_by_text(driver, "Search")
    _log(None, "Opened Account Checker search dialog.")

    _switch_to_details(driver)
    modal = _wait_element(driver, By.NAME, "myframe__1")
    driver.switch_to.frame(modal)
    account_input = _wait_element(driver, By.ID, "SEARCH_ACA_AC_NO")
    account_input.clear()
    account_input.send_keys(str(account_number).strip())
    _log(None, "Entered account number in Account Checker search.")
    _click_button_by_text(driver, "OK")
    _log(None, "Submitted Account Checker search.")
    driver.switch_to.default_content()

    _switch_to_details(driver)
    try:
        _wait_element(driver, By.CSS_SELECTOR, "form[name='fmDetails'] table", timeout=8)
    except TimeoutException:
        _log(None, "Account Checker result table did not appear.")
        _click_refresh_if_present(driver)
        return False
    rows = driver.find_elements(By.CSS_SELECTOR, "form[name='fmDetails'] table tr")
    if len(rows) <= 2:
        _log(None, "Account Checker search returned no record row.")
        _click_refresh_if_present(driver)
        return False
    links = rows[2].find_elements(By.CSS_SELECTOR, "td:nth-child(2) a")
    if not links:
        _log(None, "Account Checker result link not found.")
        _click_refresh_if_present(driver)
        return False
    links[0].click()
    _log(None, "Opened Account Checker result record.")
    _switch_to_details(driver)
    return True


def _is_red_css_color(color):
    normalized_color = normalize(color).replace(" ", "")
    return normalized_color in {"RED", "#FF0000", "RGB(255,0,0)", "RGBA(255,0,0,1)"}


def _read_account_checker_record(driver, expected_values):
    _log(None, f"Reading Account Checker record; expected values={expected_values!r}.")
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
            _log(
                None,
                f"Account Checker mismatch for {column_name}: expected={expected!r}, "
                f"actual={current[field]!r}.",
            )
            mismatches.append(column_name)
        elif normalize(expected):
            _log(None, f"Account Checker {column_name} matches ({current[field]!r}).")
        else:
            _log(None, f"Account Checker {column_name} is blank in workbook; ignored.")

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
        _log(None, f"Account Checker mismatches found: {mismatches!r}.")
        return ", ".join(mismatches)
    recommendation = "Approve" if approve_data else "Undo"
    _log(None, f"Account Checker recommendation is {recommendation}; manual action required.")
    return f"No mismatches - {recommendation} manually"


# -----------------------------------------------------------------------------
# Row processors
# These are the actual per-row automation actions called by the top-level workflow.
# Each function receives the Excel row values and executes a single maker/checker step.
# -----------------------------------------------------------------------------
def _run_checker_row(driver, values):
    _log(None, f"Customer Checker row values: {values!r}.")
    rel_id, cif = values[:2]
    expected_values = {
        "change_status": values[2],
        "branch_code": values[3],
        "officer_code": values[4],
        "customer_segment": values[5],
    }
    if not _open_checker_record(driver, cif, rel_id):
        return "CIF/Rel_ID Not Found"
    result = _read_checker_record(driver, expected_values)
    if result:
        return result
    _approve_customer_checker_record(driver)
    return "Approved"


def _run_account_checker_row(driver, values):
    _log(None, f"Account Checker row values: {values!r}.")
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
    first_row = 5 if is_account else 6
    result_column, start_column, end_column = (5, 6, 7) if is_account else (7, 8, 9)
    processor = {
        "maker": _run_maker_row,
        "checker": _run_checker_row,
        "account_maker": _run_account_maker_row,
        "account_checker": _run_account_checker_row,
    }[mode]
    stop_event = stop_event or Event()
    log_token = _WORKFLOW_LOG_CALLBACK.set(log_callback)

    try:
        _log(None, f"Starting UTS workflow mode={mode!r}, workbook={str(path)!r}.")
        _switch_to_uts_window(driver)
        for row_number in range(first_row, worksheet.max_row + 1):
            cif = worksheet.cell(row_number, 2).value
            rel_id = worksheet.cell(row_number, 1).value
            if cif in (None, "") and (is_account or is_maker or rel_id in (None, "")):
                break
            if worksheet.cell(row_number, result_column).value not in (None, ""):
                _log(log_callback, f"Row {row_number}: skipped (already has a result).")
                continue
            if stop_event.is_set():
                _log(log_callback, "Stop requested; leaving remaining rows untouched.")
                break

            _log(None, f"Starting row {row_number}; identifier={cif!r}, Rel_ID={rel_id!r}.")
            worksheet.cell(row_number, start_column, datetime.now())
            try:
                if is_account:
                    columns_to_read = range(2, 6)
                elif is_maker:
                    columns_to_read = range(2, 7)
                else:
                    columns_to_read = range(1, 7)
                values = [worksheet.cell(row_number, column).value for column in columns_to_read]
                _log(None, f"Row {row_number} input values: {values!r}.")
                result = processor(driver, values)
                worksheet.cell(row_number, result_column, result)
                _log(log_callback, f"Row {row_number}: {result}")
            except Exception as error:
                message = str(error)
                result = message if message == "ERROR_segment out of scope" else f"Error: {message}"
                worksheet.cell(row_number, result_column, result)
                _log(log_callback, f"Row {row_number}: {result}")
                if message != "ERROR_segment out of scope":
                    _log(
                        log_callback,
                        f"Row {row_number} exception traceback:\n{traceback.format_exc()}",
                    )
            finally:
                worksheet.cell(row_number, end_column, datetime.now())
                workbook.save(path)
                _log(None, f"Saved workbook after row {row_number}.")
    finally:
        workbook.close()
        _WORKFLOW_LOG_CALLBACK.reset(log_token)


def run_customer_workflow(driver, workbook_path, mode, log_callback=None, stop_event=None):
    mode = str(mode).strip().lower()
    if mode not in {"maker", "checker"}:
        raise ValueError("mode must be 'maker' or 'checker'")
    return run_uts_workflow(
        driver, workbook_path, mode, log_callback=log_callback, stop_event=stop_event
    )
