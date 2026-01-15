# PROGRESS — cl-excel

This file is a running dev log of what was implemented, how it was verified, and what comes next.
Entries are reverse-chronological (newest first).








### 2026-01-15 — Task: Enhanced write-xlsx Arguments (Milestone: M12)
**Goal:** Provide granular control over where data is written (sheet, cell, region).
**Summary:**
- Enhanced `write-xlsx` to support `:sheet` (name or index), `:start-cell`, `:region`, and `:region-only-p`.
- Implemented data clipping logic when `:region-only-p` is true.
- Enabled "best-effort" update mode when `:overwrite-p` is nil.
- Standardized testing environment using `Makefile` and Roswell (`scripts/test.ros`).
- Fixed invalid `workbook` initialization arguments.
- Improved `close-xlsx` to safely handle workbooks without an associated ZIP file.
- Updated existing tests for compatibility with system FiveAM (fixing `is` macro usage).
- Removed `vendor/` directory in favor of system-installed FiveAM.
**Files changed:**
- `src/writer.lisp`: Logic for new arguments and clipping.
- `src/workbook-read.lisp`: Moved `sheet` accessor, updated `close-xlsx`.
- `cl-excel.asd`: Added new tests.
- `Makefile`, `scripts/test.ros`, `scripts/repl.ros`: New testing infrastructure.
- `tests/write-enhancements.lisp`: New comprehensive test suite.
- `tests/run.lisp`, `tests/workbook.lisp`, `tests/tables.lisp`: Test fixes and integration.
**Commands run / verification:**
- `make test` (Executes all 119 checks across all suites).
- `ros scripts/test.ros` (Direct execution of the test runner).
**Result:** PASS


### 2026-01-14 — Refactor: Smart Ranges & Aliases (Milestone: M11)
**Goal:** Align "Sugar" API with user intuition.
**Summary:**
- Added alias `cell` for `val`/`[]`.
- Refactored `resolve-smart-range`:
    - Integers (`1`) and Numeric Strings (`"1"`) now refer to **Rows** (trimmed to data).
    - Added Row Range support: `"1:5"`.
    - Added Column Range support: `"A:C"`.
- Updated `README.md` and fixed regression tests.
**Files changed:**
- `src/sugar.lisp`: Logic updates.
- `src/package.lisp`: Exports.
- `README.md`: Documentation.
**Verification:**
- `test-cell.lisp` and `test-smart-rows.lisp`.
**Result:** PASS

### 2026-01-14 — Refactor: Robust Resource Management (Milestone: M7/M10)
**Goal:** Ensure `with-xlsx` automatically releases resources.
**Summary:**
- Refactored `with-xlsx` to use `unwind-protect` and verify `close-xlsx` is called.
- Updated documentation to reflect that explicit closing is no longer needed within `with-` blocks.
- Added demo utilities `list-examples`, `example-path`.
**Files changed:**
- `src/workbook-read.lisp`: Refactored macro.
- `src/sugar.lisp`: Added demo utils.
- `README.md`: Updated examples.
**Verification:**
- `test-with-xlsx.lisp` logic test.
**Result:** PASS

### 2026-01-14 — Task: Refactor Test Fixtures
**Goal:** Organize test files and ensure reliable path access.
**Summary:**
- Moved all Excel files (`edited.xlsx`, `original.xlsx`, `smart.xlsx`, `sugar.xlsx`, `test_table.xlsx`) to `tests/fixtures/`.
- Implemented `fixture-path` helper in `tests/run.lisp` using `asdf:system-relative-pathname`.
- Updated all test files and `README.md` to use the new fixture path location.
**Files changed:**
- `tests/run.lisp`: Added `fixture-path`.
- `tests/workbook.lisp`, `test-lazy.lisp`, `README.md`: Updated paths.
**Commands run / verification:**
- `(asdf:test-system :cl-excel)`
- `sbcl --non-interactive --load test-lazy.lisp`
**Result:** PASS

### 2026-01-14 — Bug Fix: Lazy Iterator Values (Milestone: M7)
**Goal:** Fix issue where `make-sheet-iterator` dropped cell values.
**Summary:**
- Identified that `make-sheet-iterator` parsed `<v>` content but failed to assign it to `raw-val`.
- Added missing assignment `(setf raw-val txt)` in `src/sheet-read.lisp`.
- Updated `test-lazy.lisp` to capture and print full rows effectively.
**Files changed:**
- `src/sheet-read.lisp`: Fixed value assignment.
- `test-lazy.lisp`: Added robust Quicklisp loading.
**Verification:**
- `sbcl --script test-lazy.lisp` now prints full rows (including numbers/booleans).
**Result:** PASS


### 2026-01-13 — Task: M11 Intelligent DSL (Milestone: M11)
**Goal:** Make the Sugar API "brain friendly" with intelligent range resolution.
**Summary:**
- Implemented `resolve-smart-range`:
  - Handles "A" (Column A, auto-trimmed to data height).
  - Handles "1" or integer 1 (Column A, auto-trimmed).
  - Handles "A1" (Single cell range).
- Added `list-sheets` (works on path or workbook) and `used-range` (bounding box).
- Updated `read-file` to use smart ranges.
**Files changed:**
- `src/sugar.lisp`: Added logic.
- `src/package.lisp`: Exported new symbols.
**Commands run / verification:**
- `ros -Q run --load test-smart.lisp`
**Result:** PASS

### 2026-01-13 — Task: M10 Sugar API (Milestone: M10)
**Goal:** Create a user-friendly, high-level API.
**Summary:**
- Implemented `src/sugar.lisp`.
- Added `read-file` for one-shot data reading (with range support).
- Added aliases `read-excel`, `save-excel`, `val`, `[]`.
- Added `map-rows` utility.
**Files changed:**
- `src/sugar.lisp`: New file.
- `src/iterators.lisp`: Added `do-rows` implementation.
- `src/package.lisp`: Exported new symbols.
**Commands run / verification:**
- `ros -Q run --load test-sugar.lisp`
**Result:** PASS

### 2026-01-13 — Task: M9 Edit Mode (Milestone: M9)
**Goal:** Enable modifying existing XLSX files (read-modify-write).
**Summary:**
- Implemented `open-xlsx` with `:mode :read|:write|:rw`.
- Updated `write-xlsx` and workbook/relationships writers to support dynamic multi-sheet generation from the in-memory workbook model.
- Added "best-effort" warning for `:rw` mode (unsupported parts are dropped).
**Files changed:**
- `src/workbook-read.lisp`: Added `open-xlsx`, updated `with-xlsx`.
- `src/workbook-write.lisp`: Made `write-workbook-xml` and relations dynamic.
- `src/writer.lisp`: Updated orchestration to pass sheet objects.
**Commands run / verification:**
- `ros -Q run --load test-edit.lisp`
**Result:** PASS

### 2026-01-13 — Task: M8 Writing Tables (Milestone: M8)
**Goal:** Support creating and writing Excel Tables (`ListObject`).
**Summary:**
- Implemented `add-table!` API to Create tables from data.
- Implemented `write-table-xml` to generate `tableN.xml`.
- Integrated table registration into `[Content_Types].xml` and `.rels`.
**Files changed:**
- `src/table-write.lisp`: New file.
- `src/core.lisp`: Added `add-table!`.
- `src/workbook-write.lisp`: Updated Content Types.
- `src/writer.lisp`: Updated `write-xlsx` to collect and write tables.
**Commands run / verification:**
- `ros -Q run --load test-table-write.lisp`
**Result:** PASS


### 2025-12-31 — Task: M0 skeleton + minimal tests (Milestone: M0)
**Goal:** Establish a buildable ASDF skeleton for cl-excel with a minimal passing test suite.
**Summary:**
- Added ASDF systems `:cl-excel` and `:cl-excel/tests`, and wired `test-op` so `asdf:test-system :cl-excel` runs tests.
- Created `src/package.lisp` defining package `cl-excel` (nickname `xlsx`) and exporting the full public API symbol list (stubs for future milestones).
- Added minimal `src/core.lisp` with `*implementation-version*` placeholder.
- Added FiveAM-based test harness in `tests/run.lisp` verifying package existence + exported symbols.
- Added a minimal vendored FiveAM shim under `vendor/fiveam/` to avoid external dependency fetching (temporary; replaceable with real FiveAM later).

**Files changed:**
- `cl-excel.asd`
- `src/package.lisp`
- `src/core.lisp`
- `tests/run.lisp`
- `vendor/fiveam/fiveam.asd`
- `vendor/fiveam/src/fiveam.lisp`

**Commands run / verification:**
- `ros -Q run --eval '(require :asdf)' --eval '(asdf:test-system :cl-excel)' --eval '(quit)'`
  - (If this fails because FiveAM is vendored and not on ASDF registry, run the SBCL command below or add the vendor path to ASDF registry.)
- `env HOME="$PWD" XDG_CACHE_HOME="$PWD/.cache" sbcl --no-userinit --no-sysinit --non-interactive --eval '(require :asdf)' --eval '(pushnew #p"$PWD/" asdf:*central-registry* :test #\'equal)' --eval '(pushnew #p"$PWD/vendor/fiveam/" asdf:*central-registry* :test #\'equal)' --eval '(asdf:test-system :cl-excel)'`

**Result:** PASS

**Notes / open questions:**
- Vendored FiveAM shim is intentionally minimal; decide later whether to keep for portability or switch to real FiveAM via Quicklisp.

**Next:** M1 — implement CellRef/Range parsing + `+missing+` semantics with thorough unit tests.
