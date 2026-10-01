# Customer Script to Workflow/Presets Conversion Instructions

## Goal

Convert the web-automation behavior of the supplied Excel VBA account-level and customer-level modules into this application's workflow and preset JSON format. Keep each source procedure as its own process:

- `accountlevel`: account maker workflow.
- `checkeracc`: account checker workflow.
- `sitechecker`: customer maker workflow.
- `checkerpart`: customer checker workflow.

Treat these as separate workflows. Do not combine their steps or outcomes. Convert browser interactions only; Excel row iteration, timestamps, result-cell writes, dialogs, and VBA control flow are not automatically part of the workflow engine.

## Read These References First

- `workflow.json` shows the top-level workflow list and step objects.
- `presets.json` shows the selector hierarchy and selector definitions.
- `core/workflow.py` defines the workflow data model.
- `core/automation_runner.py` defines executable actions and parameters.
- `gui/step_editor.py` defines the actions exposed in the workflow editor.
- The supplied `accountlevel` and `custLevel` VBA modules are the behavior to convert.

Match the repository's current schema, not a generic Selenium or RPA schema. Keep unrelated existing workflows and presets intact when integrating generated entries.

## Required JSON Shapes

A workflow is one object in the top-level JSON array:

```json
{
  "name": "Account Level Maker",
  "team": "Account Level",
  "screen": "Account Maintenance",
  "tab": "Search",
  "steps": [
    {
      "action": "input",
      "selector_key": "Account Level/Account Maintenance/Search/account_number",
      "params": {
        "context_key": "Account No"
      },
      "auto_generated": false
    }
  ]
}
```

A selector key must have four slash-separated parts: `team/screen/tab/field`. The same key must resolve to a preset at the corresponding nested path. Use the actual team, screen, and tab names discovered for this UTS page; the names above are illustrative only.

A preset group has this shape:

```json
{
  "Account Level": {
    "Account Maintenance": {
      "Search": {
        "account_number": {
          "type": "css",
          "selector": "VERIFIED_CSS_SELECTOR"
        }
      }
    }
  }
}
```

Use a real, tested CSS selector for every preset entry. Do not put `VERIFIED_CSS_SELECTOR` or guessed selectors into production JSON. Existing presets may also carry metadata such as `table_index`, `label_keywords`, or dynamic-table configuration where the runner supports it; only include metadata that has been verified and is relevant.

Every ordinary workflow step should include `action`, `selector_key`, `params`, and `auto_generated`, following `workflow.json`. Selectorless actions such as `switch_to_new_window` and `close_window_and_return` use `selector_key: null` and their supported parameters. Do not invent action names.

## Conversion Method

1. Read the complete VBA procedures and identify every browser operation, its page/frame context, the data it reads or enters, and the condition that controls it.
2. Inspect the actual UTS page/DOM or an approved selector source to resolve every CSS selector. VBA IDs/classes and frame indexes are clues, not proof that an equivalent Selenium selector works.
3. Add selector presets grouped by verified team, screen, tab, and field. Give fields stable descriptive names, for example `account_number`, `search_button`, `account_result`, `account_status`, `branch_code`, `edit_button`, `save_button`, `checker_search_button`, `pending_record_link`, `approve_button`, and `undo_button`.
4. Add separate workflow definitions for maker and checker. Order actions as the page requires and use `wait_for` on a verified element after navigation or submission instead of translating fixed VBA sleeps literally.
5. Map each `input` step's `params.context_key` to the exact key supplied by the application's Excel/context integration. Do not assume VBA column numbers such as B, C, and D are valid context keys. If that integration has not been confirmed, record the needed key mapping as unresolved.
6. For reads, use `extract` with `params.target_column` to name the result. For example, extract live status and branch into stable result keys before any comparison can be done by supported application logic. Do not assume an `extract` step itself performs a comparison or writes a status to the source workbook.
7. Validate that every workflow `selector_key` resolves to an existing preset, every action is implemented in `core/automation_runner.py` and exposed where needed by `gui/step_editor.py`, and each JSON file parses successfully.
8. Report unresolved selectors and unsupported behavior separately. Never silently drop or approximate a business rule.

## Confirmed Application Gaps for `custLevel`

The current runner/editor does not expose enough behavior to safely reproduce all customer workflows as JSON alone. Treat the following as required application features (runner changes and, where noted, editor support), not as invented workflow actions:

- **Nested frame navigation:** `switch_to_frame` resets to default content and locates a frame from there. The customer VBA traverses the top-level frame, a child frame, and then `myframe__1`. The existing action cannot traverse that hierarchy by issuing multiple frame steps because each step resets the context. Extend the action to accept and traverse an ordered frame path, or add an equivalent nested-frame navigation mechanism. Add Selenium tests for each frame depth and for returning to the intended parent/default context.
- **Select dropdown option:** the customer maker selects a status option by its displayed text in `CUST_DETAIL_CUST_STATUS_CODE`. `input` only clears an element and sends keys; there is no explicit select action. Add and expose an action such as `select_option` with a selector key, a confirmed option value or visible-text parameter (including a context key for dynamic values), and normal Selenium select/change event behavior. Do not assume `input` is equivalent.
- **Compare values and branch safely:** the customer maker compares extracted live values with row inputs; the checker compares only non-empty requested values and builds a mismatch list. Current `conditions` are checked only by extraction actions; click/input steps cannot be guarded, and there is no comparison action that produces reusable boolean/mismatch outputs. Add a compare/evaluate action or equivalent application logic, plus conditional execution for mutating steps based on input and extracted values. It must support trimmed, case-insensitive equality and optional/blank expected values. Test that matching values prevent edits and that only mismatching, non-empty fields are changed.
- **Read label styling for checker decision:** `checkerpart` inspects the label text and font color to derive `approveData`. `extract` reads element text, not computed style/attributes. If the business rule is confirmed to depend on red labels, add a supported attribute/computed-style extraction capability and make its result available to the comparison/branch mechanism. Do not infer style from a guessed selector.
- **Row outcome/writeback:** the macros write `Error`, `CIF Not Found`, mismatch fields, `Done`, `Approved`, or `Undo`, and timestamps to Excel columns. Confirm the application has row-level status and timestamp writeback and can skip completed rows; otherwise this requires separate Excel-runner/UI support and is not accomplished by workflow JSON.

The frame, selection, and branching gaps are independently blocking for faithful, safe execution of the customer flows. A generated JSON file must not use these proposed action names until the corresponding implementation and editor registration exist.

## Customer Workflows: `sitechecker` and `checkerpart`

### Customer Maker: `sitechecker`

Preserve the observed sequence and data mapping:

1. Search by CIF from VBA column B, entering it in `SEARCH_TEXT` and clicking the search control; then open the matching result `SPAN_CUST_DETAIL_LIST_CUST_NO_0` and wait for details.
2. Read current status (`SPAN_CUST_DETAIL_CUST_STATUS_CODE`), branch (`SPAN_CUST_DETAIL_GEN_BRCH_CODE`), and officer (`SPAN_CUST_DETAIL_GEN_OFFCR_CODE`) from the details view. Compare trimmed, case-insensitive values against desired status in column C, branch in D, and officer in E.
3. If all three match, perform no edit. Otherwise open the edit view. Change the status only when the requested status is non-empty, using the dropdown-selection feature above; set branch and officer only when their respective requested values are non-empty; then save.
4. The macro writes `Error`/`Done` to column F and start/end times to G/H. The UI interaction is not a substitute for confirming these row writebacks.

### Customer Checker: `checkerpart`

Preserve the observed sequence and data mapping:

1. Search by CIF from column B, open the matching record, and report `CIF Not Found` if the expected result is absent. The source locates a link at row 3, cell 2 of `fmDetails`; verify a stable selector or table-based lookup rather than assuming that index is invariant.
2. Read live status, branch, and officer from the nested `myframe__1` content, and compare against non-empty requested values from columns C, D, and E. Normalize by trimming and comparing case-insensitively. If there are mismatches, record the specific mismatch fields, close the detail view, and do not approve or undo.
3. When there are no mismatches, close the detail view, then apply the verified approval decision: select `MarkDel` and click Approve or Undo. The condition and chosen action must be represented by the compare/branch feature; never include both clicks as unconditional steps.
4. The macro writes `CIF Not Found`, mismatch fields, `Approved`, or `Undo` to column F and start/end times to G/H. Confirm row result persistence and skip/resume behavior separately.

**Source logic requires clarification:** in `checkerpart`, `hasOfficerCode` is initialized to `True` before scanning labels, and `approveData` is set when any of the status/branch/officer label flags is true. As written, the approval condition is therefore always true and the Undo path is unreachable. Do not silently preserve this likely initialization bug or assume the intended rule; ask the business owner to confirm the intended flag initialization and approval criterion before implementing that branch.

## Maker Workflow: `accountlevel`

Preserve this observed sequence and intent:

1. Search for the account using the account number (VBA reads column B) through `SEARCH_TEXT` and the search control.
2. Select the matching account result (`SPAN_ACCOUNT_LIST_ACA_AC_NO_0` in the VBA), then wait for the account details view.
3. Read the displayed customer/account status (`SPAN_ACCOUNT_DETAIL_CUST_STATUS_CODE`) and branch (`SPAN_ACCOUNT_DETAIL_GEN_BRCH_CODE`). Normalize comparisons by trimming whitespace and comparing case-insensitively, as the VBA does.
4. If both current values already match the requested status (column C) and branch (column D), make no change.
5. If a status change is requested, open the appropriate status-change control and select the requested status. If the existing status is `CLOSED`, the VBA first changes it to `DMI`, confirms, then performs the requested status change. Preserve this two-stage business rule exactly; verify the correct controls/options and whether the page state must be refreshed between stages.
6. If a branch code is requested, open edit, set the branch field (`ACCOUNT_DETAIL_GEN_BRCH_CODE`), and save.
7. The VBA writes `Done`/`Error` to column E and timestamps to F/G. These are Excel-processing responsibilities, not browser selectors. State clearly whether the application already supports these writebacks; otherwise flag them as requiring runner/UI implementation beyond JSON.

The maker's conditional no-op and status-update logic must not be rendered as an unconditional sequence of clicks. The workflow runner currently executes a linear list of steps; its supported `conditions` are evaluated for extraction operations, not as general guards around `click` or `input` steps. If the required branches cannot be implemented by existing application logic, deliver only the safely expressible search/read workflow and explicitly identify the required runner extension. Do not create a workflow that changes already-correct accounts or selects `DMI` for every account.

## Checker Workflow: `checkeracc`

Preserve this observed sequence and intent:

1. Skip rows already marked in VBA column E (row-processing behavior; not a browser step unless the application's Excel runner already implements it).
2. Open the checker search screen, invoke Search, enter the account number from VBA column B in the `SEARCH_ACA_AC_NO` modal input, and confirm.
3. Locate the pending record. The VBA clicks the link in the third row/second cell of `fmDetails`; do not assume fixed row/cell position is stable. Prefer a verified semantic selector or supported table extraction. If no record is found, the VBA marks `Account Not Found` and continues.
4. Open the pending record detail and read status and branch from `myframe__1` / `dummy` using the same status and branch IDs noted above. Also inspect the `tablabel th` labels and red font markers that the VBA uses to set `approveData`.
5. Compare only non-empty requested values from VBA columns C (status) and D (branch), case-insensitively after trimming. Preserve the specific mismatch names (`Change_Status`, `Branch_Code`) where the app's output mechanism supports them.
6. When there are no mismatches, follow the VBA's final decision: mark `MarkDel`, then use the Approve control if `approveData` is true, otherwise the Undo control. Confirm that the label/color rule really determines approval in the live application before encoding it.
7. The VBA writes `Account Not Found`, mismatch text, `Approved`, or `Undo` to column E and start/end timestamps to F/G. Treat these as Excel integration behavior and report gaps if the current app cannot perform them.

Checker approval/undo is conditional business logic. Do not add both Approve and Undo clicks to the same unconditional workflow. The current workflow format has no general conditional branch/step action; extraction conditions do not guard arbitrary clicks. If no existing application layer implements this decision, mark checker decision/approval as unsupported by JSON alone and specify the minimal required runner enhancement. Do not approve or undo based only on a guessed selector or an unverified interpretation of the VBA's font-color rule.

## Selector and Frame Cautions

The VBA uses IE automation, nested `<frame>` elements, and a named iframe. The Python runner uses Selenium and resolves a preset's CSS selector before `switch_to_frame`; verify the selector against the live page and check whether the target is a frame or iframe. The current frame action cannot traverse nested frames (see the required feature above). Reacquire the correct browsing context after page transitions as needed. A selector that works in the top-level document may not work inside a nested frame. Do not carry over IE DOM object references, VBA `getElementsBy...` indexes, `Application.Wait`, or `SendKeys` calls as though they were JSON actions.

Prefer stable IDs/names and specific CSS selectors. Any indexed table/link selection from the VBA must be validated against representative pages and account states. The existing `presets.json` currently has no account-level selector group, so selectors for this conversion must be discovered and verified before the generated JSON can be considered executable.

## Output and Acceptance Criteria

Deliver:

- Four workflow entries, one for each source procedure, using only actions supported by the current runner unless an extension is explicitly proposed and implemented.
- The corresponding account-level and customer-level preset group(s), with all selectors verified against the target UI.
- A concise mapping table from VBA control/data field to preset key and workflow step.
- A separate list of behavior that needs Excel/UI/runner support beyond the JSON schema, especially nested frame traversal, dropdown selection, conditional comparisons/updates, checker approval/undo, missing-record outcomes, result writes, and timestamps.

Before calling the conversion complete, confirm:

- Both JSON documents parse and preserve unrelated existing data.
- Every referenced selector key resolves to a preset with a tested CSS selector.
- Inputs use confirmed context keys and extracts use stable output names.
- Maker no-op and `CLOSED` to `DMI` behavior cannot run unsafely as unconditional actions.
- Checker mismatch and approve/undo decisions cannot fall through to an unsafe unconditional click.
- Customer maker does not edit when values already match and updates only non-empty requested fields.
- Customer checker mismatch handling cannot approve/undo, and its approval rule is confirmed rather than copied from the `hasOfficerCode = True` source bug.
- Nested frames and the customer status dropdown are supported by implemented runner actions before corresponding steps are included.
- Any unsupported behavior is explicitly reported; no VBA business rule was silently omitted.
