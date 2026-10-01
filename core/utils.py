# core/utils.py
import sys
import os
import subprocess
import re

def resource_path(relative_path: str) -> str:
    """Resolve path for resources like JSON configs.
    - If frozen (exe): exe folder
    - If dev: project root (folder containing the main script)"""
    if getattr(sys, 'frozen', False):  # Running as PyInstaller exe
        base_path = os.path.dirname(sys.executable)
    else:  # Running in dev mode
        base_path = os.path.dirname(os.path.abspath(sys.argv[0]))
    return os.path.join(base_path, relative_path)

# core/utils.py
def bundled_resource_path(relative_path):
    """For bundled resources (msedgedriver.exe, assets, etc.)."""
    if getattr(sys, 'frozen', False):  # exe mode
        base_path = sys._MEIPASS
    else:  # dev mode
        base_path = os.path.dirname(__file__)
    return os.path.join(base_path, relative_path)

def get_edge_driver_version(driver_path):
    try:
        result = subprocess.run(
            [driver_path, "--version"],
            capture_output=True,
            text=True,
            check=True
        )
        version_str = result.stdout.strip()

        # Remove anything in parentheses e.g. "(64-bit)" or "(random string)"
        version_str = re.sub(r"\s*\(.*?\)", "", version_str)

        return version_str
    except Exception as e:
        return f"Unknown WebDriver version ({e})"


def passes_all_conditions(source_data, conditions, results=None):
    results = results or {}
    print(f"[extract] Checking conditions: {conditions}")
    print(f"[context] source_data = {source_data}")

    for cond in conditions:
        field = cond.get("field")
        op = cond.get("op", "equals")
        value = cond.get("value")
        slice_spec = cond.get("slice")
        ref_key = cond.get("ref")
        
        # Handle existence checks early
        if op == "not_exists":
            val = source_data.get(field)
            if val is None or str(val).strip().lower() in ["", "-", "na", "n/a"]:
                print(f"[condition] not_exists passed: '{field}' is empty or missing (value: '{val}')")
                continue
            else:
                print(f"[condition] not_exists failed: '{field}' exists with value: '{val}'")
                print(f"[fail] Condition failed: {cond}, actual value: {val}")  # ✅ Use `val` not `actual`
                return False

        elif op == "exists":
            val = source_data.get(field)
            if val is None or str(val).strip().lower() in ["", "-", "na", "n/a"]:
                print(f"[fail] Condition failed: {cond}, actual value: {val}")
                return False
            continue


        actual = source_data.get(field, "")
        if isinstance(actual, str):
            actual = actual.strip().lower()

        # Optional slicing before comparison
        if isinstance(slice_spec, list) and len(slice_spec) == 2:
            actual = actual[slice_spec[0]:slice_spec[1]]

        # Optional value from another context field
        if ref_key:
            value = results.get(ref_key, "")
            if isinstance(value, str):
                value = value.strip().lower()

        # Normalize comparison value
        if isinstance(value, str):
            value = value.strip().lower()
        elif isinstance(value, list):
            value = [str(v).strip().lower() for v in value]

        # Apply operation
        if op == "equals":
            if actual != value:
                print(f"[fail] Condition failed: {cond}, actual value: {actual}")
                return False
        elif op == "in":
            if actual not in value:
                print(f"[fail] Condition failed: {cond}, actual value: {actual}")
                return False
        elif op == "not_in":
            if actual in value:
                print(f"[fail] Condition failed: {cond}, actual value: {actual}")
                return False
        elif op == "startswith":
            if isinstance(value, list):
                if not any(actual.startswith(v) for v in value):
                    print(f"[fail] Condition failed: {cond}, actual value: {actual}")
                    return False
            else:
                if not actual.startswith(value):
                    print(f"[fail] Condition failed: {cond}, actual value: {actual}")
                    return False
        elif op == "contains":
            if isinstance(value, list):
                if not any(re.search(rf"\b{re.escape(v)}\b", actual) for v in value):
                    print(f"[fail] Condition failed: {cond}, actual value: {actual}")
                    return False
            else:
                if not re.search(rf"\b{re.escape(value)}\b", actual):
                    print(f"[fail] Condition failed: {cond}, actual value: {actual}")
                    return False
        else:
            print(f"[warn] Unknown condition op: {op}")
            return False

    return True

