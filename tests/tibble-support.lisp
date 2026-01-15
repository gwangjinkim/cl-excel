;;;; tests/tibble-support.lisp

(in-package #:cl-excel.tests)

(def-suite :cl-excel/tibble
  :description "Test cl-tibble integration via as-tabular protocol.")
(in-suite :cl-excel/tibble)

(test write-tibble
  "Test that a cl-tibble:tbl can be written directly to XLSX."
  (let ((path (fixture-path "tibble_out.xlsx")))
    (when (probe-file path) (delete-file path))
    
    ;; Only run if cl-tibble is available
    (if (find-package :cl-tibble)
        (let ((df (funcall (find-symbol "TIBBLE" :cl-tibble)
                           :name #("Alice" "Bob")
                           :age  #(30 25))))
          (cl-excel:write-xlsx df path)
          
          ;; Verify result
          (cl-excel:with-xlsx (wb path)
            (let ((sh (cl-excel:sheet wb 1)))
              (is (equal "name" (cl-excel:get-data sh "A1")))
              (is (equal "age" (cl-excel:get-data sh "B1")))
              (is (equal "Alice" (cl-excel:get-data sh "A2")))
              (is (equal 30 (cl-excel:get-data sh "B2")))
              (is (equal "Bob" (cl-excel:get-data sh "A3")))
              (is (equal 25 (cl-excel:get-data sh "B3")))))
          
          (when (probe-file path) (delete-file path)))
        (skip "cl-tibble package not found, skipping integration test."))))
