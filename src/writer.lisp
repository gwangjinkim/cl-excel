;;;; src/writer.lisp

(in-package #:cl-excel)

(defgeneric as-tabular (source)
  (:documentation "Normalize SOURCE (e.g. list of lists, tibble) into a list of lists for writing to XLSX."))

(defmethod as-tabular ((source list))
  source)

;; Fallback to check for cl-tibble objects without a hard dependency
(defmethod as-tabular (source)
  (let* ((package (find-package :cl-tibble))
         (class (and package (find-class (find-symbol "TBL" package) nil))))
    (if (and class (typep source class))
        (let* ((names (coerce (funcall (find-symbol "TBL-NAMES" package) source) 'list))
               (nrows (funcall (find-symbol "TBL-NROWS" package) source))
               (get-col (find-symbol "TBL-COL" package))
               (rows (loop for i from 0 below nrows
                           collect (loop for name in names
                                         collect (aref (funcall get-col source name) i)))))
          (cons names rows))
        (error "Unsupported source type: ~A. Consider implementing CL-EXCEL:AS-TABULAR." (type-of source)))))



(defun write-entry-xml (zip name callback)
  "Write XML to a ZIP entry by buffering it first (ZIP library limitation)."
  (let* ((content-str (with-output-to-string (s) (funcall callback s)))
         (content-bytes (flexi-streams:string-to-octets content-str :external-format :utf-8)))
    (let ((in (flexi-streams:make-in-memory-input-stream content-bytes)))
      (zip:write-zipentry zip name in :file-write-date (get-universal-time)))))

(defun write-sheet-rels-xml (sheet stream tables-start-id)
  "Write xl/worksheets/_rels/sheetN.xml.rels"
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "Relationships"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/package/2006/relationships")
      (loop for tbl in (sheet-tables sheet)
            for i from 1
            for global-id = (+ tables-start-id i -1)
            do (cxml:with-element "Relationship"
                 (cxml:attribute "Id" (format nil "rId~D" i))
                 (cxml:attribute "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table")
                 (cxml:attribute "Target" (format nil "../tables/table~D.xml" global-id)))))))

(defun write-xlsx (source path &key sheet start-cell region region-only-p (overwrite-p t))
  "Write data to an XLSX file at PATH.
   SOURCE can be a workbook object or tabular data (normalized via as-tabular).
   SHEET: name or 1-based index (default 1).
   START-CELL: e.g. \"A1\".
   REGION: e.g. \"A1:D10\".
   REGION-ONLY-P: if T, clip data to fit REGION.
   OVERWRITE-P: if NIL and PATH exists, edit the existing file."
  (ensure-directories-exist path)
  
  (let ((wb (cond
              ;; Case 1: SOURCE is already a workbook
              ((typep source 'workbook) source)
              
              ;; Case 2: SOURCE is tabular data (normalized)
              (t
               (let ((data (as-tabular source)))
                 (let ((existing-wb (when (and (not overwrite-p) (probe-file path))
                                      (handler-case (read-xlsx path)
                                        (error () nil)))))
                   (let ((wb (or existing-wb (make-instance 'workbook :zip nil :sheets nil))))
                     ;; Identify target sheet
                     (let* ((target-sheet-id (or sheet 1))
                            (sh (handler-case (sheet wb target-sheet-id)
                                  (sheet-missing-error ()
                                    ;; Create new sheet if missing
                                    (let ((new-sh (make-instance 'sheet 
                                                                 :name (if (stringp target-sheet-id) 
                                                                           target-sheet-id 
                                                                           (format nil "Sheet~D" target-sheet-id))
                                                                 :id (1+ (length (workbook-sheets wb))))))
                                      (setf (workbook-sheets wb) (append (workbook-sheets wb) (list new-sh)))
                                      new-sh)))))
                       
                       ;; Resolve writing window (start-r, start-c, end-r, end-c)
                       (let ((start-r 1) (start-c 1) (end-r nil) (end-c nil))
                         ;; Region takes precedence
                         (if region
                             (let ((r (parse-cell-ref region :range t)))
                               (setf start-r (range-start-row r)
                                     start-c (range-start-col r)
                                     end-r (range-end-row r)
                                     end-c (range-end-col r)))
                             (when start-cell
                               (let ((c (parse-cell-ref start-cell)))
                                 (setf start-r (cell-ref-row c)
                                       start-c (cell-ref-col c)))))
                         
                         ;; Write data
                         (loop for row in data
                               for r from start-r
                               while (or (not region-only-p) (not end-r) (<= r end-r))
                               do (loop for val in row
                                        for c from start-c
                                        while (or (not region-only-p) (not end-c) (<= c end-c))
                                        do (setf (get-cell sh (cons r c))
                                                 (make-cell val))))))
                     wb)))))))
    
    ;; Final write to disk
    
    ;; Collect all tables to assign global filenames
    (let ((all-tables '())
          (table-counter 1))
      (dolist (sh (workbook-sheets wb))
        (dolist (tbl (sheet-tables sh))
          (push (cons tbl table-counter) all-tables)
          (incf table-counter)))
      (setf all-tables (nreverse all-tables))
      
      (zip:with-output-to-zipfile (zip path :if-exists :supersede)
        ;; 1. [Content_Types].xml
        (write-entry-xml zip "[Content_Types].xml" 
                         (lambda (s) 
                           (write-content-types-xml 
                            s 
                            :sheets (workbook-sheets wb)
                            :tables (loop for i from 1 below table-counter collect (format nil "table~D" i)))))
        
        ;; 2. _rels/.rels
        (write-entry-xml zip "_rels/.rels" 
                         (lambda (s) (write-rels-xml s)))
        
        ;; 3. xl/workbook.xml
        (write-entry-xml zip "xl/workbook.xml" 
                         (lambda (s) (write-workbook-xml s :sheets (workbook-sheets wb))))
        
        ;; 4. xl/_rels/workbook.xml.rels
        (write-entry-xml zip "xl/_rels/workbook.xml.rels" 
                         (lambda (s) (write-workbook-rels-xml s :sheets (workbook-sheets wb))))
        
        ;; 5. xl/styles.xml
        (write-entry-xml zip "xl/styles.xml" 
                         (lambda (s) (write-styles-xml s)))
        
        ;; 6. Sheets and their rels
        (let ((global-tbl-idx 1))
          (dolist (sh (workbook-sheets wb))
            (let ((sh-path (format nil "xl/worksheets/sheet~D.xml" (sheet-id sh)))
                  (sh-rels-path (format nil "xl/worksheets/_rels/sheet~D.xml.rels" (sheet-id sh))))
              
              (write-entry-xml zip sh-path (lambda (s) (write-sheet-xml sh s)))
              
              (when (sheet-tables sh)
                (write-entry-xml zip sh-rels-path 
                                 (lambda (s) (write-sheet-rels-xml sh s global-tbl-idx)))
                (incf global-tbl-idx (length (sheet-tables sh)))))))
        
        ;; 7. Tables
        (loop for (tbl . id) in all-tables
              do (write-entry-xml zip (format nil "xl/tables/table~D.xml" id)
                                  (lambda (s) (write-table-xml tbl s))))))))
