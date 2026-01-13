# PROGRESS — cl-excel

This file is a running dev log of what was implemented, how it was verified, and what comes next.
Entries are reverse-chronological (newest first).










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
