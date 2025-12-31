;;;; Minimal local FiveAM compatibility shim for offline tests.

(asdf:defsystem #:fiveam
  :description "Tiny FiveAM-compatible test shim (local vendor copy for cl-excel M0)."
  :version "0.0.0"
  :author "cl-excel contributors"
  :license "MIT"
  :serial t
  :components
  ((:module "src"
    :serial t
    :components ((:file "fiveam")))))

