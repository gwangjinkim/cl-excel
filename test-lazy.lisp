
(require :asdf)

;; Ensure Quicklisp is loaded for dependencies (Zip, CXML, etc.)
(unless (find-package :ql)
  (let ((ql-setup (merge-pathnames "quicklisp/setup.lisp" (user-homedir-pathname))))
    (if (probe-file ql-setup)
        (load ql-setup)
        (format *error-output* "Warning: Quicklisp not found at ~A. Dependencies might fail.~%" ql-setup))))

(push (uiop:getcwd) asdf:*central-registry*)
(asdf:load-system :cl-excel)

(defun test-lazy ()
  (format t "Testing Lazy Iterator...~%")
  (let ((wb (cl-excel:read-xlsx (asdf:system-relative-pathname :cl-excel "tests/fixtures/basic_types.xlsx"))))
    (unwind-protect
         (cl-excel:with-sheet-iterator (next-row wb "Sheet1")
           (loop for row = (funcall next-row)
                 while row
                 do (format t "Row: ~S~%" row)))
      (cl-excel:close-xlsx wb))))

(test-lazy)
(sb-ext:exit)
