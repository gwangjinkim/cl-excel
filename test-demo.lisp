
(require :asdf)

;; Ensure Quicklisp is loaded
(unless (find-package :ql)
  (let ((ql-setup (merge-pathnames "quicklisp/setup.lisp" (user-homedir-pathname))))
    (if (probe-file ql-setup)
        (load ql-setup)
        (format *error-output* "Warning: Quicklisp not found. Dependencies might fail.~%"))))

(push (uiop:getcwd) asdf:*central-registry*)
(asdf:load-system :cl-excel)

(format t "Testing Demo Utilities...~%")

(let ((path (cl-excel:example-path "basic_types.xlsx")))
  (format t "Example Path: ~A~%" path)
  (unless (probe-file path)
    (error "Example file not found!")))

(let ((examples (cl-excel:list-examples)))
  (format t "Found ~D examples:~%" (length examples))
  (dolist (ex examples)
    (format t " - ~A~%" (file-namestring ex)))
  (unless (> (length examples) 0)
    (error "No examples found!")))

(sb-ext:exit)
