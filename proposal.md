# Proposal: Feature List for Unified UTS Automation Application

## 1. Business Objective

The new application will combine the existing account-level and customer-level workflows into one Python-based automation tool using Selenium. The goal is to replace the current duplicated Excel/VBA process with a single, cleaner, and more scalable system for processing UTS data from Excel.

The application will support the following core capabilities:

- customer-level validation and update processing
- account-level validation and update processing
- Excel-driven batch processing
- live comparison between Excel inputs and UTS record values
- status, branch, and officer updates
- error logging, screenshot capture, and retry handling
- row-level progress and resume support

---

## 2. Feature List

### 2.1 Unified Workflow Engine

One application should handle both customer and account processing from the same interface.

Features:

- select workflow type: customer or account
- select Excel file and target sheet
- process all rows or only selected rows
- skip completed rows automatically
- resume from failed rows without restarting the whole batch
- operate with a common processing flow for both record types

### 2.2 Excel-Based Batch Processing

The application should read input data from Excel and process each row in sequence.

Supported fields may include:

- CIF number
- account number
- desired status
- branch code
- officer code
- result/status column
- start time column
- end time column
- notes or mismatch column

Features:

- dynamic mapping of Excel columns
- validation of required values before processing
- automatic status updates back to Excel
- timestamp logging for each processed row
- processing summary at the end of the batch

### 2.3 Search and Record Discovery

The app should navigate to the UTS system and locate the target record by the values provided in Excel.

Features:

- customer search by CIF
- account search by account number
- result validation after search
- detection of missing or invalid records
- direct navigation to the correct detail page
- support for repeated retries if the page is slow or partially loaded

### 2.4 Data Comparison and Validation

After opening the record, the app should compare Excel-based expected values with the live values displayed on the page.

Features:

- compare status values
- compare branch values
- compare officer values
- classify each row as approved, mismatch, error, or missing
- show the specific mismatched field(s)
- prevent edits when no valid record is found

### 2.5 Update Execution

When the live values do not match the expected values, the application should perform the required update in the UTS system.

Features:

- open the correct edit screen
- update customer or account status when required
- populate branch code and officer code values
- click save/confirm actions
- verify the update was applied successfully
- record final result back to Excel

### 2.6 Retry and Resilience

The system should be designed to handle unreliable page loads and changing DOM states.

Features:

- retry stale element lookups
- wait for dynamic content to load before acting
- re-fetch frame and iframe documents after refreshes
- retry when buttons or forms are temporarily unavailable
- handle missing elements without crashing the full batch
- continue processing after recoverable failures

### 2.7 Error Handling and Audit Logging

The application should capture operational issues in a clear and traceable way.

Features:

- row number and record identifier in logs
- workflow name and action step
- timestamped exception entries
- screenshot capture for failed steps
- final status classification for each row
- log file and summary report generation

### 2.8 Resume and Batch Control

The application should make it possible to resume large jobs safely.

Features:

- start from a specific row
- continue from failed entries only
- skip rows already marked complete
- stop gracefully on user request
- run in dry-run mode for validation before actual updates

### 2.9 Reporting and Status Tracking

The tool should provide clear results after processing.

Features:

- summary of total rows processed
- count of successful rows
- count of mismatches
- count of failed rows
- count of skipped rows
- per-row result output in Excel
- optional export to CSV or report file

### 2.10 User-Friendly Interface

The application should be easy for staff to operate without needing to edit code.

Features:

- select workbook and worksheet
- choose workflow type
- choose batch mode or resume mode
- view live progress
- display processing logs in real time
- show error summary at the end of run

---

## 3. Proposed High-Level Modules

### 3.1 Core Application Layer

- main entry point
- configuration management
- workflow selection
- execution control
- logging setup

### 3.2 Excel Processing Layer

- file loading
- column mapping
- validation rules
- row status writing
- output reporting

### 3.3 Browser Automation Layer

- browser startup
- session management
- page navigation
- frame and iframe handling
- explicit waits and element interaction

### 3.4 UTS Workflow Layer

- customer search workflow
- account search workflow
- live value retrieval
- comparison logic
- update execution logic

### 3.5 Error and Reporting Layer

- screenshot capture
- exception logging
- summary output
- retry statistics
- final status report

---

## 4. Expected Benefits

This Python + Selenium application will provide:

- one unified system for both customer and account processing
- less duplication than the current VBA scripts
- easier maintenance and future updates
- clearer user experience and workflow control
- more reliable handling of page timing and dynamic content
- better traceability of errors and processing outcomes

---

## 5. Recommended Next Step

The next phase should focus on converting these feature requirements into a technical design and implementation plan, including:

1. field mapping for Excel inputs
2. workflow-specific rules for customer and account processing
3. Selenium page object structure
4. retry and wait strategy
5. reporting and row-status design

That will give us a practical blueprint for building the application in Python.

The current VBA automation works by directly interacting with IE-mode frames, nested documents, and multiple hard-coded page states. While it has been adapted over time, several structural issues make it risky to continue using as-is.

### 2.1 Repeated Logic Across Modules

Both scripts repeat similar patterns for:

- opening browser sessions
- fetching frames and documents
- waiting for page load
- reading nested frames and iframes
- searching by CIF or account number
- clicking buttons and navigating to detail pages
- editing status, branch, and officer values

This duplication increases maintenance cost and makes future changes harder to apply consistently.

### 2.2 Fragile Frame and DOM Handling

The automation depends on a number of assumptions:

- frame indexes like `allFrames(1)` or `innerframe(0)`
- fixed button positions (`getElementsByClassName(...)(0)`)
- static text matching for buttons
- nested DOM access that is not validated before use

These assumptions break easily when the page layout changes or when the browser loads slower or differently.

### 2.3 Weak Error Handling

There are scattered `On Error Resume Next` patterns, silent exits, and `GoTo` logic that makes the process hard to trace. Some cases record generic status values but do not explain the actual failure cause.

The system needs structured errors with:

- step name
- row number
- CIF/account number
- screenshot capture
- retry count
- reason for failure

### 2.4 Hard-Coded Timers and Timing Issues

The scripts rely on fixed waits such as `Application.Wait Now + TimeSerial(0, 0, 3)`. These do not account for actual page readiness and can create race conditions during slow network or application response states.

The new solution should use explicit waits for:

- element visibility
- element clickability
- iframe availability
- document readiness
- AJAX load completion

### 2.5 Inconsistent State Tracking

Current implementation often marks rows as `Done`, `Error`, or leaves them blank depending on the branch of logic. This makes auditing and retrying a challenge.

The new system should standardize row statuses such as:

- Pending
- In Progress
- Done
- Error
- Skipped
- Validation Failed

### 2.6 Unclear Data Validation Logic

In several places, the scripts compare Excel values and live values but do not clearly separate:

- data mismatch detection
- approval conditions
- close/refresh actions
- row outcome classification

The rebuilt app should clearly define validation rules before any update is submitted.

### 2.7 Browser Dependency and Session Stability

The current automation depends on Edge IE mode and specific browser behavior. This creates a brittle dependency on legacy compatibility. The rebuild should standardize the browser setup and make session management much more robust.

---

## 3. Proposed Solution

Build a single Python application that centralizes common workflow logic and handles both account-level and customer-level processing through a shared automation engine.

### 3.1 Core Architecture

The application can be structured into the following layers:

- UI layer: configuration, workbook selection, run controls, progress logs
- Business logic layer: row processing, validation, workflow rules
- Automation layer: Selenium browser driver, page objects, waits, selectors
- Data layer: Excel I/O, row status tracking, logs, reports
- Error handling layer: retries, exceptions, screenshot capture, audit records

### 3.2 Recommended Modules

#### a) App Core
- main entry point
- configuration management
- logger setup
- environment and browser settings
- run mode selection (single row, batch, resume)

#### b) Excel Processor
- read input rows from workbook
- map columns to expected values
- validate required fields
- update outputs back to Excel
- track timestamps and status

#### c) Browser Manager
- open browser session
- attach to existing browser if available
- create stable waits and session management
- handle page refresh and frame recovery

#### d) UTS Page Objects
- login and navigation
- search page
- employee/customer page
- account detail page
- edit/update page
- frame/iframe utility methods

#### e) Workflow Engine
- account workflow
- customer workflow
- comparison and validation functions
- update execution rules
- close/refresh handling

#### f) Monitoring and Logging
- row-by-row task log
- failure reasons
- screenshot capture on error
- summary report after run

---

## 4. Functional Features to Include

### 4.1 Unified Processing Engine

One application should support both customer-level and account-level flows without duplicating logic.

Features:

- select workflow type: customer or account
- select workbook and sheet dynamically
- process all rows or only selected rows
- resume from failed rows
- skip rows already completed

### 4.2 Excel-Driven Batch Processing

The application should read rows from Excel and process them in sequence using mapped columns such as:

- CIF number
- account number
- desired status
- branch code
- officer code
- expected result column
- timestamp columns
- status column

This keeps data entry and processing fully configurable.

### 4.3 Search and Validation Workflow

The app should handle the search flow for:

- CIF number lookup
- account number lookup
- result validation
- live page data extraction
- comparing expected values vs live values

### 4.4 Dynamic Data Comparison

The system should compare the expected values from Excel against the actual values on the UTS page and classify results as:

- Approved
- Mismatch
- Error
- Missing record
- Invalid selector

### 4.5 Update Execution

When values differ, the app should:

- navigate to the correct edit screen
- update status, branch, and officer values as provided
- click save or confirm actions
- verify that the update took effect
- record final status in Excel

### 4.6 Retry and Resilience

The application should implement a structured and reusable retry mechanism, including:

- retrying stale element lookups
- re-fetching frame documents after refresh
- waiting for dynamic content to render
- page reload fallback
- safe handling of timeout and missing elements

### 4.7 Exception Management

Each failure should be recorded with enough detail to troubleshoot later:

- row number
- workflow name
- step name
- browser state
- screenshot path
- timestamp
- exception message

### 4.8 Resume Capability

The app should be able to continue from the last successful or failed row without restarting the full workbook.

### 4.9 Reporting and Audit Trail

At the end of each execution, the app should produce:

- summary counts
- processed rows
- successful rows
- failed rows
- mismatched rows
- skipped rows
- error log file
- optional CSV or Excel report

### 4.10 User-Friendly Interface

A simple desktop UI should help non-technical users run the process safely:

- browse Excel file
- select sheet
- choose workflow
- configure start row
- toggle pause/resume
- view live logs
- stop execution cleanly

---

## 5. Technical Fixes and Improvements Required

### 5.1 Replace Hard-Coded DOM Indexing

Current VBA code uses many static indexes, which is highly brittle. The Python version should replace this with:

- explicit element IDs where possible
- name-based lookup
- XPath fallback
- CSS selectors with validation
- data attributes if available

This reduces failures when page layouts or lists change.

### 5.2 Centralize Repeated Frame Logic

The current scripts repeat the same frame-fetching pattern many times. The Python app should introduce a single frame manager that:

- locates the main content frame
- re-fetches frames after navigation
- validates nested element availability
- handles iframe switching safely
- returns the active document for the current workflow

### 5.3 Use Explicit Waits Instead of Static Timers

Instead of arbitrary sleep loops, the Python app should use Selenium WebDriverWait and expected conditions such as:

- presence_of_element_located
- visibility_of_element_located
- frame_to_be_available_and_switch_to_it
- element_to_be_clickable
- staleness_of

This will make the automation more reliable and faster.

### 5.4 Remove Fragile `GoTo` and Exit Logic

The current VBA module uses jump logic and early exits that make behavior difficult to track. The Python rebuild should use structured control flow and clear methods such as:

- process_row()
- validate_page()
- update_record()
- handle_exception()
- record_result()

This creates a more maintainable and testable workflow.

### 5.5 Standardize State and Outcome Handling

A consistent row status model should be implemented, for example:

- Pending
- Running
- Completed
- Mismatch
- Error
- Skipped

This makes reporting clearer and allows safer resumption.

### 5.6 Improve Data Validation Before Actions

Before performing any update, the app should validate:

- the record exists
- the expected page is loaded
- the required values are present
- the field names match current selectors
- the input values are not empty or malformed

This prevents accidental updates based on stale or partial data.

### 5.7 Add Screenshot Capture and Replay Support

On failure, the app should automatically capture screenshots of:

- current browser page
- active frame or modal
- error step context

This makes debugging far easier and supports future replay and root cause analysis.

### 5.8 Separate Search, Read, and Update Logic

The existing script mixes search, comparison, and update actions in one flow. The Python version should clearly separate these tasks into independent steps.

Recommended separation:

- locate record
- read live values
- compare with expected values
- decide if update is required
- execute update
- verify result

This reduces complexity and improves reliability.

### 5.9 Add Configuration and Environment Controls

The application should support configuration for:

- browser type and profile
- page URL
- timeout settings
- retry counts
- Excel column mappings
- log folder location
- screenshot folder location

This makes the system easier to deploy and maintain in different environments.

### 5.10 Add Safe Batch Execution Controls

The app should include controls for:

- dry run mode
- test single row mode
- stop on first error mode
- continue on error mode
- email or notification output after completion

---

## 6. Recommended Python Design Patterns

To make the application robust and maintainable, the following patterns are recommended:

### 6.1 Page Object Model (POM)

Each UTS page should have a dedicated page class with methods like:

- open_page()
- search_by_cif()
- search_by_account()
- open_detail_record()
- compare_values()
- update_record()

This keeps selectors isolated from business logic.

### 6.2 Service Layer

A service layer can manage the actual workflows, such as:

- customer_workflow()
- account_workflow()
- compare_record_data()
- update_status_fields()

### 6.3 Utility Layer

A utilities layer can include reusable functions for:

- Excel read/write
- wait logic
- retry loops
- logging
- screenshot capture

### 6.4 Data Model Layer

The application should use simple data structures for row records, for example:

- row_id
- cif_number
- account_number
- requested_status
- branch_code
- officer_code
- status_output
- timestamps

This improves clarity and makes testing easier.

---

## 7. Proposed Deliverables

The final application should include:

- Python desktop application or CLI runner
- Selenium browser automation layer
- Excel import/export processing
- data validation and comparison functions
- logging and reporting module
- failure screenshot capture
- documentation for setup and usage

---

## 8. Expected Benefits

The new Python and Selenium solution will provide the following benefits:

- one application for both account and customer flows
- easier troubleshooting and maintenance
- better handling of dynamic pages and nested frames
- improved error visibility and audit logs
- reduced duplication of VBA logic
- more scalable processing for future changes
- better support for batch execution and resume functions

---

## 9. Recommended Next Steps

1. Map the exact fields and business rules from the current VBA logic.
2. Define the standard data structure for input/output Excel files.
3. Choose the Selenium browser configuration and target platform.
4. Build the shared page-object and workflow framework.
5. Validate customer workflow against the live system.
6. Validate account workflow.
7. Add logging, screenshots, and resume support.
8. Run regression tests against sample batches.

---

## 10. Final Recommendation

The rebuild should not be a direct translation of the current VBA scripts. Instead, it should be treated as a redesign based on the same business process, but with a cleaner architecture, stronger automation standards, and better operational control.

The key principle is to replace brittle automation logic with a maintainable, data-driven system built on Selenium and Python. This will reduce operational risk and make future enhancements much easier to implement.
