;;;; src/package.lisp — packages and exports (M0)

(defpackage #:cl-excel
  (:use #:cl)
  (:nicknames #:excel)
  (:export
   ;; Open/close
   #:read-xlsx #:open-xlsx #:close-xlsx #:with-xlsx #:write-xlsx
   ;; Workbook/sheets
   #:sheet-names #:sheet-count #:has-sheet-p #:sheet #:add-sheet! #:rename-sheet! #:sheet-name #:sheet-cells
   ;; Cells/ranges
   #:get-data #:read-data #:get-cell #:get-cell-range #:cell-value
   #:+missing+ #:missing-p
   ;; Internal utils often needed
   #:col-name #:col-index #:parse-cell-ref #:cell-ref-to-string
   ;; Indexing sugar (optional API surface)
   #:cell #:range #:sheet-ref #:named
   ;; Iterators
   #:each-row #:row-number #:do-rows #:each-table-row #:do-table-rows
   #:make-sheet-iterator #:with-sheet-iterator ;; M7 exports
   ;; Tables
   #:get-table #:read-table #:add-table! 
   #:table #:table-name #:table-ref #:table-columns #:table-display-name
   #:table-column #:table-column-name #:table-column-id
   ;; Conditions
   #:xlsx-error #:xlsx-parse-error #:sheet-missing-error
   #:invalid-range-error #:read-only-error
   ;; Sugar (M10/M11)
   #:read-excel #:save-excel #:sheet-of #:val #:[] #:c 
   #:map-rows #:with-sheet #:read-file
   #:list-sheets #:used-range #:with-xlsx-save
   #:list-examples #:example-path))

