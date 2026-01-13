;;;; src/iterators.lisp (M7)

(in-package #:cl-excel)

;;; Helper to find sheet-meta
(defun find-sheet-meta (workbook name-or-index)
  ;; We need access to the metas which are not currently exposed on workbook struct easily?
  ;; Workbook struct has `sheets` which are fully loaded SHEET objects (M2).
  ;; But wait, lazy loading implies we might NOT want to load them all?
  ;; For M7, we assume the workbook structure is loaded.
  ;; We need the meta (ID, etc) to open the XML stream again.
  ;; The `sheet` object has ID/RelID.
  (let ((sh (sheet workbook name-or-index)))
    ;; We reconstruct meta from sheet object
    (make-sheet-meta :name (sheet-name sh)
                     :id (sheet-id sh)
                     :rel-id (sheet-rel-id sh))))

(defmacro with-sheet-iterator ((iter-var workbook sheet-designator) &body body)
  "Execute BODY with ITER-VAR bound to a function `(next-row)`."
  (let ((wb-sym (gensym "WB"))
        (sh-sym (gensym "SH"))
        (cleanup-sym (gensym "CLEANUP")))
    `(let* ((,wb-sym ,workbook)
            (,sh-sym (find-sheet-meta ,wb-sym ,sheet-designator)))
       (multiple-value-bind (,iter-var ,cleanup-sym)
           (make-sheet-iterator (workbook-zip ,wb-sym) 
                                ,sh-sym 
                                (workbook-shared-strings ,wb-sym)
                                (workbook-styles ,wb-sym))
         (unwind-protect
              (progn ,@body)
           (funcall ,cleanup-sym))))))
