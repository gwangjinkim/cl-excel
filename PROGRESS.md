# PROGRESS — cl-excel

This file is a running dev log of what was implemented, how it was verified, and what comes next.
Entries are reverse-chronological (newest first).











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
