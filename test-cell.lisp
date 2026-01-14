
(require :asdf)

;; Ensure Quicklisp is loaded
(unless (find-package :ql)
  (let ((ql-setup (merge-pathnames "quicklisp/setup.lisp" (user-homedir-pathname))))
    (if (probe-file ql-setup)
        (load ql-setup)
        (format *error-output* "Warning: Quicklisp not found. Dependencies might fail.~%"))))

(push (uiop:getcwd) asdf:*central-registry*)
(asdf:load-system :cl-excel)

(format t "Testing 'cell' alias...~%")

(let ((path (cl-excel:example-path "basic_types.xlsx")))
  (cl-excel:with-xlsx (wb path :mode :rw)
    (cl-excel:with-sheet (sh wb 1)
      
      ;; Test Reading
      (let ((val (cl-excel:cell sh "A2")))
        (format t "Read A2 via 'cell': ~A~%" val)
        (unless (equal val "Hello")
          (error "Failed to read using 'cell'")))
      
      ;; Test Writing (setf cell)
      (setf (cl-excel:cell sh "A2") "Changed")
      (let ((new-val (cl-excel:cell sh "A2")))
        (format t "Read A2 after write: ~A~%" new-val)
        (unless (equal new-val "Changed")
          (error "Failed to write using '(setf cell)'"))))))

(format t "Success!~%")
(sb-ext:exit)
