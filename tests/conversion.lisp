;;;; tests/conversion.lisp

(in-package #:cl-excel.tests)

(def-suite :cl-excel/conversion
  :description "Test high-level conversion functions (LOL, ALIST, PLIST, TIBBLE).")
(in-suite :cl-excel/conversion)

(test excel-to-lol-test
  "Test conversion to list-of-lists."
  (let ((path (fixture-path "basic_types.xlsx")))
    (let ((data (cl-excel:excel-to-lol path)))
      (is (listp data))
      (is (listp (first data)))
      ;; basic_types.xlsx has "String", "Integer", "Float", "Bool", "Date", "DateTime", "Missing"
      (is (equal "String" (first (first data))))
      (is (equal "Integer" (second (first data)))))))

(test excel-to-alist-test
  "Test conversion to list of alists."
  (let ((path (fixture-path "basic_types.xlsx")))
    (let ((data (cl-excel:excel-to-alist path)))
      (is (listp data))
      (let ((first-row (first data)))
        (is (listp first-row))
        ;; Check if headers are used as keys
        (is (assoc "String" first-row :test #'string=))
        (is (assoc "Integer" first-row :test #'string=))
        (is (equal "Hello" (cdr (assoc "String" first-row :test #'string=))))))))

(test excel-to-plist-test
  "Test conversion to list of plists."
  (let ((path (fixture-path "basic_types.xlsx")))
    (let ((data (cl-excel:excel-to-plist path)))
      (is (listp data))
      (let ((first-row (first data)))
        (is (listp first-row))
        ;; Check if headers are used as keyword keys
        (is (equal "Hello" (getf first-row :STRING)))
        (is (equal 42 (getf first-row :INTEGER)))))))

(test excel-to-tibble-test
  "Test conversion to tibble (if available)."
  (if (find-package :cl-tibble)
      (let ((path (fixture-path "basic_types.xlsx")))
        (let ((data (cl-excel:excel-to-tibble path)))
          ;; Check if it's a tibble (duck typing or check type if possible)
          (is (not (null data)))
          (is (string-equal "tbl" (package-name (symbol-package (type-of data)))))))
      (skip "cl-tibble not found, skipping excel-to-tibble test.")))
