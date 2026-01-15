;;;; cl-excel.asd — ASDF system definition (M0 skeleton)

;;;; Keep this file minimal for Milestone M0. Subsequent milestones will
;;;; extend the component list and dependencies.

(asdf:defsystem #:cl-excel
  :description "Common Lisp XLSX reader/writer."
  :version "0.1.0"
  :author "cl-excel contributors"
  :license "MIT"
  :depends-on (#:zip #:cxml #:cxml-stp #:local-time #:flexi-streams #:parse-number)
  :serial t
  :in-order-to ((test-op (test-op "cl-excel/tests")))
  :components
  ((:module "src"
    :serial t
    :components
    ((:file "package")
     (:file "types")
     (:file "refs")
     (:file "zip")
     (:file "xml")
     (:file "rels")
     (:file "styles")
     (:file "sheet-read")
     (:file "table-read")
     (:file "table-write")
     (:file "workbook-read")
     (:file "sheet-write")
     (:file "workbook-write")
     (:file "writer")
     (:file "iterators")
     (:file "core")
     (:file "sugar")))))

(asdf:defsystem #:cl-excel/tests
  :description "Test suite for cl-excel (M0)."
  :depends-on (#:cl-excel #:fiveam)
  :serial t
  :components
  ((:module "tests"
    :serial t
    :components ((:file "run")
                 (:file "refs")
                 (:file "workbook")
                 (:file "tables")
                 (:file "write-enhancements")
                 (:file "tibble-support"))))
  :perform (test-op (op c)
             (uiop:symbol-call '#:cl-excel.tests '#:run-tests)))
