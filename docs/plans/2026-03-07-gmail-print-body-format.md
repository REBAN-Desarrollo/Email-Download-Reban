# Gmail Print Body Format Implementation Plan

> **For Claude:** REQUIRED SUB-SKILL: Use superpowers:executing-plans to implement this plan task-by-task.

**Goal:** Add a selectable email body output mode that can save Gmail-style PDFs, original HTML, or both.

**Architecture:** Move body rendering decisions into a dedicated helper module. Use the raw email HTML only for the original artifact, and generate the Gmail PDF from a sanitized fragment mounted into a fixed Gmail-style print template.

**Tech Stack:** Python 3, tkinter, IMAP/email stdlib, Playwright worker process, `unittest`

---

### Task 1: Add test coverage for the new rendering helpers

**Files:**
- Create: `tests/test_email_rendering.py`
- Test: `tests/test_email_rendering.py`

**Step 1: Write the failing test**

Create tests that define the new contract:
- sanitization removes document wrappers and unsafe print/layout styles
- Gmail print document includes the expected metadata shell
- original HTML document remains standalone and readable

**Step 2: Run test to verify it fails**

Run: `py -3 -m unittest tests.test_email_rendering -v`
Expected: FAIL because `email_rendering.py` does not exist yet.

**Step 3: Write minimal implementation**

Create `email_rendering.py` with:
- body format constants/options
- HTML extraction/sanitization helpers
- Gmail print document builder
- original HTML document builder

**Step 4: Run test to verify it passes**

Run: `py -3 -m unittest tests.test_email_rendering -v`
Expected: PASS

### Task 2: Connect the new rendering helpers to the download flow

**Files:**
- Modify: `download_handler.py`
- Modify: `pdf_worker.py`

**Step 1: Write the failing test**

Use the rendering helper tests as the behavior guard for the HTML transformation contract before wiring the UI.

**Step 2: Run test to verify it fails if helper behavior regresses**

Run: `py -3 -m unittest tests.test_email_rendering -v`
Expected: PASS before refactor, then keep rerunning after download flow changes.

**Step 3: Write minimal implementation**

Update the download flow to:
- choose outputs by selected body format
- save `Mensaje_Original.html` when required
- generate `Mensaje_Gmail.pdf` from sanitized HTML when required
- save `Mensaje_Gmail.html` if PDF generation fails
- only start the worker when a PDF is requested

**Step 4: Run test to verify it passes**

Run: `py -3 -m unittest tests.test_email_rendering -v`
Expected: PASS

### Task 3: Add the body format selector to the UI and settings

**Files:**
- Modify: `app.py`
- Modify: `utils.py` (only if helper constants/settings glue belongs there)

**Step 1: Write the failing test**

Use manual verification for the tkinter selector because the repo has no GUI test harness.

**Step 2: Run manual check to verify current UI lacks the selector**

Run the app and confirm there is no `Formato del cuerpo` selector yet.

**Step 3: Write minimal implementation**

Add:
- selector UI
- enable/disable behavior linked to `Descargar cuerpo`
- settings load/save support
- mode lookup helper used by the download flow

**Step 4: Run manual verification**

Run the app and confirm:
- selector appears
- selector disables when body download is off
- selection persists after save/reload

### Task 4: Verify the real PDF flow end-to-end

**Files:**
- Modify if needed: `download_handler.py`
- Modify if needed: `email_rendering.py`

**Step 1: Run targeted verification**

Run:
- `py -3 -m unittest tests.test_email_rendering -v`
- `py -3 -c "import py_compile; py_compile.compile('app.py', doraise=True); py_compile.compile('download_handler.py', doraise=True); py_compile.compile('email_rendering.py', doraise=True); py_compile.compile('pdf_worker.py', doraise=True)"`

**Step 2: Run real PDF generation smoke test**

Generate a real Gmail-style PDF through the Playwright worker using representative HTML with inline images.

**Step 3: Inspect result**

Confirm:
- the first page starts with real content, not a giant blank region
- inline images render
- Gmail shell appears

**Step 4: Prepare for review**

Capture `git diff` and request a code review before closing the task.
