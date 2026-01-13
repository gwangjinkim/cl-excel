;;;; src/zip.lisp

(in-package #:cl-excel)

(defclass xlsx-zip ()
  ((handle :initarg :handle :accessor zip-handle)
   (entries :initarg :entries :accessor zip-entries)))

(defun open-zip (path)
  "Open a ZIP file and return an XLSX-ZIP object."
  (let ((h (zip:open-zipfile path)))
    (make-instance 'xlsx-zip 
                   :handle h
                   :entries (zip:zipfile-entries h))))

(defun close-zip (xlsx-zip)
  "Close the underlying ZIP handle."
  (zip:close-zipfile (zip-handle xlsx-zip)))

(defmacro with-zip ((var path) &body body)
  `(let ((,var (open-zip ,path)))
     (unwind-protect
          (progn ,@body)
       (close-zip ,var))))

(defun get-entry-stream (xlsx-zip name)
  "Return an input stream for the ZIP entry named NAME.
   Returns NIL if not found."
  (let ((entry (gethash name (zip-entries xlsx-zip))))
    (if entry
        (let ((contents (zip:zipfile-entry-contents entry)))
          (flexi-streams:make-in-memory-input-stream contents))
        nil)))
