;;;; src/writer.lisp

(in-package #:cl-excel)



(defun write-entry-xml (zip name callback)
  "Write XML to a ZIP entry by buffering it first (ZIP library limitation)."
  (let* ((content-str (with-output-to-string (s) (funcall callback s)))
         (content-bytes (flexi-streams:string-to-octets content-str :external-format :utf-8)))
    (let ((in (flexi-streams:make-in-memory-input-stream content-bytes)))
      (zip:write-zipentry zip name in :file-write-date (get-universal-time)))))

(defun write-xlsx (source path &key (args nil))
  "Write data to an XLSX file at PATH.
   SOURCE can be:
   - A list of lists (rows of values): Creates a single sheet workbook.
   - A WORKBOOK object: Writes the workbook structure (Not fully supported for editing yet).
  "
  (declare (ignore args))
  
  (ensure-directories-exist path)
  
  (zip:with-output-to-zipfile (zip path :if-exists :supersede)
    ;; 1. [Content_Types].xml
    (write-entry-xml zip "[Content_Types].xml" 
                     (lambda (s) (write-content-types-xml s)))
    
    ;; 2. _rels/.rels
    (write-entry-xml zip "_rels/.rels" 
                     (lambda (s) (write-rels-xml s)))
    
    ;; 3. xl/workbook.xml
    (write-entry-xml zip "xl/workbook.xml" 
                     (lambda (s) (write-workbook-xml s)))
    
    ;; 4. xl/_rels/workbook.xml.rels
    (write-entry-xml zip "xl/_rels/workbook.xml.rels" 
                     (lambda (s) (write-workbook-rels-xml s)))
      
    ;; 5. xl/styles.xml
    (write-entry-xml zip "xl/styles.xml" 
                     (lambda (s) (write-styles-xml s)))
      
    ;; 6. xl/worksheets/sheet1.xml
    (write-entry-xml zip "xl/worksheets/sheet1.xml"
                     (lambda (s)
                       ;; If list-of-lists, mock a sheet
                       (let ((sh (if (listp source)
                                     (let ((new-sh (make-instance 'sheet :name "Sheet1" :id 1)))
                                       (loop for row in source
                                             for r from 1
                                             do (loop for val in row
                                                      for c from 1
                                                      do (setf (get-cell new-sh (cons r c))
                                                               (make-cell val))))
                                       new-sh)
                                     (if (typep source 'workbook)
                                         (first (workbook-sheets source)) ;; Just write first sheet for M6 MVP
                                         (error "Unsupported source type")))))
                         (write-sheet-xml sh s))))))
