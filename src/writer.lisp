;;;; src/writer.lisp

(in-package #:cl-excel)



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

(defun write-xlsx (source path &key (args nil))
  "Write data to an XLSX file at PATH."
  (declare (ignore args))
  (ensure-directories-exist path)
  
  ;; Normalize source to Workbook
  (let ((wb (if (typep source 'workbook)
                source
                (if (listp source)
                    (let ((new-wb (make-instance 'workbook :id nil :rels nil)))
                      (let ((new-sh (make-instance 'sheet :name "Sheet1" :id 1)))
                        (loop for row in source
                              for r from 1
                              do (loop for val in row
                                       for c from 1
                                       do (setf (get-cell new-sh (cons r c))
                                                (make-cell val))))
                        (setf (workbook-sheets new-wb) (list new-sh))
                        new-wb))
                    (error "Unsupported source type")))))
    
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
                         (lambda (s) (write-workbook-xml s))) ;; TODO: Update write-workbook-xml to use real sheets!
        
        ;; 4. xl/_rels/workbook.xml.rels
        (write-entry-xml zip "xl/_rels/workbook.xml.rels" 
                         (lambda (s) (write-workbook-rels-xml s))) ;; TODO: Dynamic
        
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
