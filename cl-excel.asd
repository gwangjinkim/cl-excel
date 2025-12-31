;;;; cl-excel.asd — ASDF system definition (M0 skeleton)

;;;; Keep this file minimal for Milestone M0. Subsequent milestones will
;;;; extend the component list and dependencies.

(asdf:defsystem #:cl-excel
  :description "Common Lisp XLSX reader/writer — skeleton for M0."
  :version "0.0.1"
  :author "cl-excel contributors"
  :license "MIT"
  :serial t
  :in-order-to ((test-op (test-op "cl-excel/tests")))
  :components
  ((:module "src"
    :serial t
    :components
    ((:file "package")
     (:file "core")))))

(asdf:defsystem #:cl-excel/tests
  :description "Test suite for cl-excel (M0)."
  :depends-on (#:cl-excel #:fiveam)
  :serial t
  :components
  ((:module "tests"
    :serial t
    :components ((:file "run"))))
  :perform (test-op (op c)
             (uiop:symbol-call '#:cl-excel.tests '#:run-tests)))
