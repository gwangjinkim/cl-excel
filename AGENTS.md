# AGENTS.md — cl-excel (Common Lisp XLSX reader/writer)

## Mission
Build `cl-excel`, a Common Lisp package that reads/writes Excel **.xlsx** files with feature parity to **Julia XLSX.jl** (as documented in its Tutorial + API Reference).

Parity means:
- Same core operations (open/read/write/edit, sheet navigation, cell/range access, iterators, table read/write helpers).
- Same *data-type semantics* (typed values inferred from cell content + style where applicable).
- Similar limitations: “edit mode” (`rw`) is best-effort and may drop unknown OOXML parts; safe for simple data, risky around complex features (e.g., charts/formulas).

Non-goals:
- Legacy `.xls` support.
- Full-fidelity preservation of every OOXML part during `rw` (charts, drawings, pivot caches, etc.) unless explicitly required for parity.
- Full Excel calculation engine. We may preserve formulas as strings; cached results can be read when present.

Primary user experience goal:
> “I can read and write Excel files in Common Lisp as seamlessly as XLSX.jl does in Julia.”

---

## Product scope: XLSX.jl parity checklist

### 1) File open/read/write
Implement equivalents of:
- `readxlsx(source::path|io) -> workbook` (eager load)
- `openxlsx(source; mode="r"|"w"|"rw", enable_cache=true|false)` with:
  - **do-syntax equivalent** via a macro: `(with-xlsx (wb path :mode :r :enable-cache t) ...)`
  - non-macro API for manual close if desired
- `writexlsx(output, workbook; overwrite=false)` (write to path or stream)

Modes:
- `:r`  read
- `:w`  create new, blank workbook
- `:rw` edit existing (best-effort; warn loudly in docs)

Caching:
- `enable_cache=true`: cache read cells
- `enable_cache=false`: always read from disk/streaming parse (for huge sheets)

### 2) Workbook & sheet navigation
- `sheetnames(workbook) -> list-of-strings`
- `sheetcount(workbook) -> integer`
- `hassheet(workbook, name) -> boolean`
- Sheet access:
  - by name: `sheet(workbook, "mysheet")`
  - by 1-based index: `sheet(workbook, 1)`
- Sheet creation + rename for writing:
  - `addsheet!(workbook, "name") -> sheet`
  - `rename!(sheet, "new_name")`

### 3) Cell/range addressing (string references)
Support the same addressing ergonomics:
- single cell: `"B2"`
- rectangular range: `"A2:B4"`
- entire used range: `:all` (like `sheet[:]` in Julia)
- column ranges: `"A:B"`
- sheet-qualified: `"mysheet!A2:B4"`
- named ranges: `"NAMED_CELL"` at workbook scope

### 4) Reading values vs reading cell objects
Two layers:
- **Value layer**:
  - `getdata(sheet, ref)` returns scalar or 2D matrix of values
  - `readdata(source, sheet, ref)` convenience that opens and reads
- **Cell object layer**:
  - `getcell(sheet, "A1") -> cell-object` (EmptyCell if missing)
  - `getcellrange(sheet, "A1:B2") -> 2D matrix of cell-objects`

### 5) Iterators (streaming-friendly)
- `eachrow(sheet)` -> iterator yielding `SheetRow`
  - row number: `row-number(sheetrow)`
  - column access: `sheetrow[1]` or `sheetrow["B"]` or `getcell(sheetrow, 2)`
- Table iterators + helpers:
  - `eachtablerow(sheet, &key columns first-row column-labels header stop-in-empty-row stop-in-row-function keep-empty-rows)`
  - `gettable(sheet, &key ...) -> DataTable`
  - `readtable(source, sheet, &key ...) -> DataTable`
- Required `readtable/gettable` options parity:
  - `columns` as `"B:E"` or inferred contiguous block
  - `first_row` (first non-empty if omitted)
  - `header` default true
  - `column_labels` override header labels
  - `infer_eltypes` (typed columns vs Any)
  - `stop_in_empty_row` default true
  - `stop_in_row_function` predicate to end table
  - `keep_empty_rows` keep/drop rows of all missing

### 6) Supported value types (XLSX.jl parity)
XLSX.jl supports:
- String
- Missing
- Float64
- Int
- Bool
- Date
- Time
- DateTime

In Common Lisp, define a concrete, portable mapping:
- String => `string`
- Missing => `cl-excel:+missing+` sentinel (NOT `NIL`)
- Float64 => `double-float`
- Int => `integer`
- Bool => `t` or `nil` (but note: nil conflicts with “missing”, hence sentinel)
- Date/Time/DateTime => use `local-time` or a small internal struct suite:
  - `date` (Y-M-D)
  - `time` (H:M:S[.n])
  - `datetime` (date+time)
Decision: default to `local-time` if available; else provide internal structs and conversion functions.

Write-time conversions:
- Abstract numeric => coerce to integer or double-float
- `nil` values in user input => convert to `+missing+` (matches “Nothing -> Missing” behavior)
- Provide `(missing-p x)` predicate and `(maybe-value x)` helper.

Type inference:
- Use **cell style** to infer:
  - date stored as number + date format => return Date/DateTime
  - numeric visualized with decimals => return float even if stored as integer
- If cell is empty => return `+missing+`
- If cell contains empty string => return `+missing+`

---

## Architecture (how to implement XLSX safely)

XLSX is a zip of XML parts (OOXML / ECMA-376). Implement as:
1) ZIP container reader
2) Relationship resolver (`.rels`)
3) Workbook model builder:
   - workbook.xml
   - sharedStrings.xml
   - styles.xml (for number formats + date detection)
   - worksheets/sheetN.xml
   - definedNames (named ranges)
4) Value extraction + streaming parsers

### Recommended internal modules
- `src/package.lisp` — packages, exports
- `src/types.lisp` — workbook/sheet/cell structs, +missing+, predicates
- `src/refs.lisp` — CellRef/Range parsing and formatting (A1, A:B, A1:B2, sheet!ref)
- `src/zip.lisp` — zip abstraction (read entry bytes as stream)
- `src/xml.lisp` — XML parsing helpers (DOM for small docs, SAX for sheet streaming)
- `src/rels.lisp` — relationship resolution
- `src/workbook-read.lisp` — parse workbook + shared strings + styles + defined names
- `src/sheet-read.lisp` — read cells, getdata, getcell, iterators, caching
- `src/table.lisp` — readtable/gettable/eachtablerow implementation
- `src/workbook-write.lisp` — generate OOXML parts, shared strings, styles, workbook, rels
- `src/sheet-write.lisp` — write cells/ranges, ensure dimension, merge shared strings
- `src/edit-mode.lisp` — best-effort `rw` rules + warning banners
- `tests/` — FiveAM suites + fixtures

### Library dependencies (choose stable, portable)
- XML: `cxml` (SAX+DOM) or `plump` (DOM) + a SAX option for big sheets
- ZIP: `zip` / `archive` (pick one; ensure binary-safe streams)
- Time: `local-time` (preferred)
- Testing: `fiveam`
- Optional: `babel` for encoding, `flexi-streams`

Codex should pick one stack and be consistent across the codebase.

---

## Public API (Common Lisp naming)
Export the following symbols from package `CL-EXCEL`:

### Open/close
- `readxlsx (source) -> workbook`
- `openxlsx (source &key mode enable-cache) -> workbook`        ; manual close
- `closexlsx (workbook) -> nil`
- `with-xlsx ((wb source &key mode enable-cache) &body body)`   ; do-syntax equivalent
- `writexlsx (output workbook &key overwrite) -> output`

### Workbook/sheets
- `sheetnames (workbook)`
- `sheetcount (workbook)`
- `hassheet (workbook name)`
- `sheet (workbook which)` ; which = string name or integer index
- `addsheet! (workbook name)`
- `rename! (sheet new-name)`

### Cells/ranges
- `getdata (sheet ref)`              ; scalar or 2D
- `readdata (source sheet ref)`      ; convenience
- `getcell (sheet ref) -> cell`
- `getcellrange (sheet range) -> 2D cell matrix`

### Indexing sugar (optional but recommended)
- `(cell sheet "B2")` and `(setf (cell sheet "B2") value)`
- `(range sheet "A1:B3")` and `(setf (range sheet "A1:B3") matrix)`
- `(sheetref workbook "mysheet!A1:B2")` ; parse sheet-qualified ref
- Named ranges: `(named workbook "NAME")`

### Iterators
- `eachrow (sheet &key left right)` -> iterator of `sheetrow`
- `row-number (sheetrow|tablerow)`
- `getcell (sheetrow column)` ; overload for sheetrow
- `eachtablerow (sheet &key columns first-row column-labels header stop-in-empty-row stop-in-row-function keep-empty-rows)`

### Tables
- `readtable (source sheet &key ...) -> datatable`
- `gettable (sheet &key ...) -> datatable`
- `datatable-data (dt)` => columns or rows (define)
- `datatable-column-labels (dt)`
- `datatable-column-index (dt)` => mapping label->index
- `writetable (output &rest sheetspecs) -> output`
- `writetable! (sheet table &key anchor-cell)` ; anchor default "A1"

Table input formats supported:
- (columns, labels) style: `(:columns (list-of-vectors) :labels (list-of-strings/symbols))`
- list of plists/alists (row-wise)
- optional adapter protocol for user-defined table objects

---

## Edit mode (`:rw`) policy (match XLSX.jl warning spirit)
- Implement `:rw` as: read existing workbook + allow modifications + write out.
- Preserve as much as we can, but:
  - Unknown parts may be dropped (charts/drawings/etc).
  - Provide a *strong warning* in README and docstrings.
- Add an option `:rw-strict t` (optional):
  - If true, refuse to open `:rw` when unsupported parts are detected.

---

## Implementation constraints / quality bar
- Never silently conflate FALSE with MISSING.
  - `nil` must mean boolean false where applicable.
  - missing must be `+missing+`.
- All public functions must have docstrings and examples.
- No global mutable state; workbook holds caches.
- Streaming reading must work for huge sheets:
  - When `enable-cache=nil`, do not retain per-cell values indefinitely.
- Be conservative on writing:
  - Generate minimal OOXML that Excel can open.
  - Add more parts only when required.

---

## Tests (must-have)
Use FiveAM.

### Fixture strategy
Commit small `.xlsx` fixtures under `tests/fixtures/`:
- `basic_types.xlsx`: string/int/float/bool/missing/date/time/datetime
- `ranges.xlsx`: A1:B2, A:B column range, sheet-qualified refs
- `named_ranges.xlsx`: workbook-level defined names
- `tables.xlsx`: header + data + empty rows + stop conditions
- `big_sparse.xlsx`: sparse rows and columns

### Test categories
1) **Unit**: ref parsing (A1, A:B, A1:B2, sheet!ref), column letters<->numbers.
2) **Read**: `readxlsx`, `getdata`, `getcell/getcellrange`, named ranges.
3) **Iterators**: `eachrow` and `eachtablerow` behavior on caching on/off.
4) **Tables**: `readtable/gettable` options parity:
   - infer columns
   - first_row
   - header vs no header
   - column_labels override
   - stop_in_empty_row true/false
   - stop_in_row_function predicate
   - keep_empty_rows
5) **Write**: create new file, write cells/ranges, `writetable`/`writetable!`.
6) **Roundtrip**: write -> read -> equals (with missing semantics).
7) **Edit-mode warning**: `:rw` emits warning; basic edits persist for simple files.

---

## Codex workflow rules (how to work in this repo)
When you (Codex/agent) modify code:

1) Keep diffs small. One feature per PR-sized change.
2) Always run tests after each change (or at least before finishing the task).
3) Update docs for every new public symbol.
4) If a feature is uncertain (OOXML nuance), add a failing test first, then implement.
5) Prefer portable Common Lisp (SBCL + CCL). Avoid implementation-specific I/O tricks.

When implementing a feature, follow this order:
- Add/extend fixture if needed.
- Add test(s).
- Implement code.
- Ensure all tests pass.

- When running or reporting tests, use the **Roswell canonical test command** from “Tooling: Roswell”.
- Include the exact Roswell command(s) executed in PROGRESS.md under “Commands run / verification”.

---

## Progress logging (mandatory): PROGRESS.md

After **every** Codex task/milestone (even small ones), update `PROGRESS.md`.

Rules:
1) If `PROGRESS.md` does not exist, create it with the header template below.
2) Append a **new entry at the top** (reverse-chronological).
3) Keep entries concise and factual. No marketing text.
4) Include: what changed, where (files), how verified (tests/commands), and next steps.
5) If tests were not run, explicitly say why and what to run.

Required entry template (copy exactly):

### YYYY-MM-DD — Task: <short title> (Milestone: Mx or “misc”)
**Goal:** <one sentence>
**Summary:** <3–8 bullet points of what was implemented/fixed>
**Files changed:**
- `<path>`
- `<path>`
**Commands run / verification:**
- `<command>`
- `<command>`
**Result:** PASS | FAIL (with 1–2 lines if FAIL)
**Notes / open questions:** <optional, bullet list>
**Next:** <one concrete next step>

Also update `README.md` only when the task introduces or changes public API.

---

## Tooling: Roswell (mandatory)

All development and tests MUST be run via **Roswell**. Do not invoke `sbcl`/`ccl` directly.

### Required commands

#### Start a REPL
```bash
ros run
```

#### Load / quick sanity check
```bash
ros -Q run --eval '(require :asdf)' \
            --eval '(asdf:load-system :cl-excel)' \
            --eval '(quit)'
```

#### Run tests (the canonical command)
```bash
ros -Q run --eval '(require :asdf)' \
            --eval '(asdf:test-system :cl-excel)' \
            --eval '(quit)'
```

#### Notes

- Use -Q to avoid user init files that might affect reproducibility.
- If multiple Lisp implementations are needed, document them an drun via: `ros -L <impl> -Q run ...`
- If tests require extra enviornment variables (e.g., fixture paths), document them in PROGRESS.md.

- `ros -Q run --eval '(require :asdf)' --eval '(asdf:test-system :cl-excel)' --eval '(quit)'` passes.
- PROGRESS.md is updated with the command output summary.


---




## Milestones (execution plan)

### M0 — Skeleton & build
- ASDF system `cl-excel`
- Packages and exports
- FiveAM setup with `tests/run.lisp`
Acceptance: `(asdf:test-system :cl-excel)` runs.

### M1 — Ref parsing + missing semantics
- CellRef and Range parsing/formatting
- `+missing+`, `missing-p`, conversions
Acceptance: ref unit tests pass.

### M2 — Read-only workbook load (eager)
- ZIP open
- parse workbook.xml -> sheets
- parse sharedStrings.xml
- parse styles.xml minimal (enough for date detection)
- `readxlsx`, `sheetnames`, `sheetcount`, `hassheet`, `sheet`
Acceptance: can open fixture and list sheets.

### M3 — Cell access & ranges
- `getcell`, `getcellrange`, `getdata`
- support "A1", "A1:B2", ":all", "A:B", "sheet!ref", named ranges
Acceptance: read fixtures.

### M4 — Streaming iterators + caching
- `eachrow` (SheetRow)
- caching toggle
Acceptance: `enable-cache=nil` doesn’t grow memory on big_sparse fixture (basic sanity).

### M5 — Table API parity
- `eachtablerow`, `gettable`, `readtable` with option parity
- `infer_eltypes` behavior (typed columns)
Acceptance: table fixtures match expected.

### M6 — Writing new files
- `openxlsx :w` to create workbook
- cell/range assignment
- `writexlsx`
Acceptance: Excel opens produced file; roundtrip tests pass.

### M7 — writetable/writetable!
- accept columns+labels and row-wise inputs
- `anchor-cell` support
Acceptance: fixtures generated by tests read back correctly.

### M8 — Edit mode (`:rw`) best-effort
- open existing -> modify -> write out
- warning banner + optional strict mode
Acceptance: simple edits roundtrip; complex fixtures not guaranteed but don’t crash.

---

## Notes on OOXML details (implementation hints)
- Cells in sheet XML are sparse; missing cells default to `+missing+`.
- Handle shared strings (`t="s"` with index into sharedStrings).
- Handle inline strings (`t="inlineStr"`) if present.
- Handle booleans (`t="b"`).
- Numeric default when `t` omitted.
- Styles: number formats used to detect date/time/datetime and decimal display.
- Named ranges/defined names live in workbook structures; map name -> (sheet, range).

---

## Documentation deliverables
- README with:
  - “Quick start” read + write examples
  - `:rw` warning section
  - Table read/write examples
- API reference page in docstrings (optionally with `40ants-doc` later)

---

## License and attribution
Do not copy XLSX.jl source code. Re-implement based on OOXML standard behavior and the public API semantics described in the XLSX.jl docs.

End.
