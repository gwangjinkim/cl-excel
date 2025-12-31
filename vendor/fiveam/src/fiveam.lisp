;;;; Minimal FiveAM shim — implements only what's needed for M0 tests.

(defpackage #:fiveam
  (:use #:cl)
  (:export #:def-suite #:in-suite #:test #:is #:run!))

(in-package #:fiveam)

(defvar *suites* (make-hash-table :test 'equal))
(defvar *current-suite* nil)

(defmacro def-suite (name &key description)
  (declare (ignore description))
  `(progn
     (setf (gethash ,name *suites*) (or (gethash ,name *suites*) '()))
     ,name))

(defmacro in-suite (name)
  `(setf *current-suite* ,name))

(defmacro is (form &optional (desc nil) &rest args)
  (let ((msg (gensym "MSG")))
    `(let ((,msg ,(if desc `(format nil ,desc ,@args) "")))
       (unless ,form
         (error (if (plusp (length ,msg))
                    ,msg
                    (format nil "Assertion failed: ~S" ',form))))
       t)))

(defmacro test (name &body body)
  `(progn
     (defun ,name ()
       ,@body)
     (when (null *current-suite*)
       (error "fiveam:test used without an active in-suite."))
     (push ',name (gethash *current-suite* *suites*))
     ',name))

(defun run! (&optional suite)
  (let ((executed 0)
        (failures 0))
    (labels ((run-suite (sname)
               (dolist (fn (reverse (gethash sname *suites*)))
                 (incf executed)
                 (handler-case
                     (funcall (symbol-function fn))
                   (error (e)
                     (incf failures)
                     (format t "~&[FAIL] ~A: ~A~%" fn e))))))
      (if suite
          (run-suite suite)
          (maphash (lambda (k _v) (declare (ignore _v)) (run-suite k)) *suites*))
      (when (> failures 0)
        (error "fiveam-shim: ~D test(s) failed." failures))
      (values executed (- executed failures)))))

