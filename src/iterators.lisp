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

(defmacro do-rows ((row-var sheet) &body body)
  "Iterate over rows in SHEET (in-memory).
   ROW-VAR is bound to the list of cell values for each row."
  (let ((rows-sym (gensym "ROWS"))
        (r-sym (gensym "R"))
        (pair-sym (gensym "PAIR"))
        (coord-sym (gensym "COORD"))
        (cell-sym (gensym "CELL"))
        (row-idx-sym (gensym "ROW-IDX")))
    `(let ((,rows-sym (make-hash-table :test 'eql)))
       ;; Group cells by row index (inefficient but necessary for hash-based storage)
       (maphash (lambda (,coord-sym ,cell-sym)
                  (push (cons (cdr ,coord-sym) (cell-value ,cell-sym)) 
                        (gethash (car ,coord-sym) ,rows-sym)))
                (sheet-cells ,sheet))
       
       (dolist (,row-idx-sym (sort (alexandria:hash-table-keys ,rows-sym) #'<))
         (let* ((,pair-sym (sort (gethash ,row-idx-sym ,rows-sym) #'< :key #'car))
                ;; Extract just values, filling gaps with +missing+ if needed? 
                ;; For basic sugar, just values is fine, or simple list.
                (,row-var (mapcar #'cdr ,pair-sym)))
           ,@body)))))

(defmacro do-table-rows ((row-var table-desig &optional sheet) &body body)
  "Iterate over rows of a TABLE.
   TABLE-DESIG can be a table object or name (if SHEET provided)."
  (let ((tbl-sym (gensym "TBL")))
    `(let ((,tbl-sym (if (typep ,table-desig 'table) 
                         ,table-desig 
                         (get-table (if ,sheet ,sheet (error "Sheet required for table name lookup")) ,table-desig))))
       (dolist (,row-var (read-table (table-sheet ,tbl-sym) (table-name ,tbl-sym)))
         ,@body))))
