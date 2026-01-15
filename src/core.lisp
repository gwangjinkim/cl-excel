;;;; src/core.lisp

(in-package #:cl-excel)

(defparameter *implementation-version* "0.1.0")

;;; Workbook Access


;;; Cell Access


(defun get-cell (sheet ref)
  "Get the raw CELL struct at REF (string 'A1', CELL-REF struct, or (row . col) cons).
   Returns a CELL with value +missing+ if empty."
  (let* ((coords (typecase ref
                   (cons ref)
                   (string (let ((r (parse-cell-ref ref)))
                             (cons (cell-ref-row r) (cell-ref-col r))))
                   (cell-ref (cons (cell-ref-row ref) (cell-ref-col ref)))
                   (t (error "Invalid ref type: ~A" (type-of ref)))))
         (raw-cell (gethash coords (sheet-cells sheet))))
    (or raw-cell (make-cell +missing+))))

(defun (setf get-cell) (new-value sheet ref)
  "Set the cell at REF to NEW-VALUE (a CELL struct)."
  (let* ((coords (typecase ref
                   (cons ref)
                   (string (let ((r (parse-cell-ref ref)))
                             (cons (cell-ref-row r) (cell-ref-col r))))
                   (cell-ref (cons (cell-ref-row ref) (cell-ref-col ref)))
                   (t (error "Invalid ref type: ~A" (type-of ref))))))
    (setf (gethash coords (sheet-cells sheet)) new-value)))

(defun get-data (sheet ref)
  "Get the value of cell(s) at REF.
   REF can be 'A1', 'A1:B2'.
   Returns scalar value or 2D array."
  ;; For M3, we focus on scalar A1 access.
  ;; TODO: cell range support logic.
  (let ((cell (get-cell sheet ref)))
    (cell-value cell)))

;;; Table Access (M4)

(defun get-table (workbook name-or-index)
  "Get a TABLE object by name (string) or index (integer)."
  (let ((tables (workbook-tables workbook)))
    (typecase name-or-index
      (integer 
       (if (and (>= name-or-index 1) (<= name-or-index (length tables)))
           (nth (1- name-or-index) tables)
           nil))
      (string 
       (find name-or-index tables :key #'table-name :test #'string-equal))
      (t nil))))

(defun read-table (workbook name)
  "Read data from the named table. Returns a list of lists (rows)."
  (let ((tbl (get-table workbook name)))
    (unless tbl 
      (error 'xlsx-error :message (format nil "Table ~A not found" name)))
    
    (let* ((sheet (cl-excel::table-sheet tbl))
           (ref (cl-excel::table-ref tbl)))
      (unless sheet
        ;; If table has no sheet reference (scanning failed to link or M4 legacy),
        ;; we can't read data.
        ;; But with M5 rels parsing, it SHOULD be there.
        (error 'xlsx-error :message (format nil "Table ~A is not linked to a sheet (Rels missing?)." name)))
      
      ;; Parse range "A2:C5"
      (let* ((range (parse-cell-ref ref :range t)) ;; Returns a range struct
             (start-row (cl-excel::range-start-row range))
             (end-row (cl-excel::range-end-row range))
             (start-col (cl-excel::range-start-col range))
             (end-col (cl-excel::range-end-col range))
             (data '()))
        
        ;; Iterate rows
        (loop for r from start-row to end-row do
          (let ((row-data '()))
            (loop for c from start-col to end-col do
              (let* ((cell-ref (format nil "~A~A" (col-name c) r))
                     (val (get-data sheet cell-ref)))
                (push val row-data)))
            (push (nreverse row-data) data)))
        (nreverse data)))))


;;; Table Creation (M8)

(defun add-table! (sheet data &key (name "Table1") (display-name "Table1") (header t) (start-cell "A1"))
  "Write DATA (list of lists) to SHEET starting at START-CELL, and register an Excel Table.
   If HEADER is true, the first row of DATA is used as column names.
   Returns the new TABLE object."
  
  (let* ((start-ref (parse-cell-ref start-cell))
         (start-r (cell-ref-row start-ref))
         (start-c (cell-ref-col start-ref))
         (row-count (length data))
         (col-count (if data (length (first data)) 0))
         (end-r (+ start-r row-count -1))
         (end-c (+ start-c col-count -1)))
    
    (when (or (zerop row-count) (zerop col-count))
      (error "Cannot add empty table."))

    ;; Write data to cells
    (loop for row in data
          for r from start-r
          do (loop for val in row
                   for c from start-c
                   do (setf (get-cell sheet (cons r c)) 
                            (make-cell val))))
    
    ;; Create Columns
    (let ((columns '())
          (header-row (if header (first data) nil)))
      (loop for c from 0 below col-count
            for col-idx from start-c
            do (let* ((raw-name (if (and header (nth c header-row))
                                    (format nil "~A" (nth c header-row))
                                    (format nil "Column~D" (1+ c))))
                      (col-id (1+ c))
                      (col (make-table-column :id col-id :name raw-name)))
                 (push col columns)))
      
      ;; Create Table Object
      (let* ((ref-str (format nil "~A~D:~A~D" 
                              (col-name start-c) start-r 
                              (col-name end-c) end-r))
             ;; Generate a simplified name for 'name' attribute (no spaces)
             (safe-name (substitute #\_ #\Space name)) 
             (tbl (make-instance 'table 
                                 :id (1+ (length (sheet-tables sheet))) ;; Simple ID generation
                                 :name safe-name
                                 :display-name display-name
                                 :ref ref-str
                                 :columns (nreverse columns)
                                 :sheet sheet)))
        
        (push tbl (sheet-tables sheet))
        tbl))))
