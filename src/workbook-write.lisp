;;;; src/workbook-write.lisp

(in-package #:cl-excel)

(defun write-content-types-xml (stream &key (sheets '("sheet1")) (tables nil))
  "Write [Content_Types].xml to STREAM.
   SHEETS is a list of sheet names (e.g. 'sheet1') or objects.
   TABLES is a list of table part names (e.g. 'table1')."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "Types"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/package/2006/content-types")
      ;; Default extension overrides
      (cxml:with-element "Default" 
        (cxml:attribute "Extension" "rels") 
        (cxml:attribute "ContentType" "application/vnd.openxmlformats-package.relationships+xml"))
      (cxml:with-element "Default" 
        (cxml:attribute "Extension" "xml") 
        (cxml:attribute "ContentType" "application/xml"))
      
      ;; Specific parts
      (cxml:with-element "Override"
        (cxml:attribute "PartName" "/xl/workbook.xml")
        (cxml:attribute "ContentType" "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"))
      (cxml:with-element "Override"
        (cxml:attribute "PartName" "/xl/styles.xml")
        (cxml:attribute "ContentType" "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"))
      
      ;; Sheets
      (dolist (sh sheets)
        ;; If sh is just a string name like "sheet1"
        (let ((name (if (stringp sh) sh (format nil "sheet~D" (sheet-id sh)))))
          (cxml:with-element "Override"
            (cxml:attribute "PartName" (format nil "/xl/worksheets/~A.xml" name))
            (cxml:attribute "ContentType" "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"))))
      
      ;; Tables
      (dolist (tbl tables)
        (cxml:with-element "Override"
          (cxml:attribute "PartName" (format nil "/xl/tables/~A.xml" tbl))
          (cxml:attribute "ContentType" "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"))))))

(defun write-rels-xml (stream)
  "Write _rels/.rels to STREAM."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "Relationships"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/package/2006/relationships")
      (cxml:with-element "Relationship"
        (cxml:attribute "Id" "rId1")
        (cxml:attribute "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        (cxml:attribute "Target" "xl/workbook.xml")))))


(defun write-workbook-rels-xml (stream &key sheets)
  "Write xl/_rels/workbook.xml.rels to STREAM."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "Relationships"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/package/2006/relationships")
      ;; Sheets
      (loop for sh in sheets
            for i from 1
            do (let ((name (if (stringp sh) sh (format nil "sheet~D" (sheet-id sh)))))
                 (cxml:with-element "Relationship"
                   (cxml:attribute "Id" (format nil "rId~D" i))
                   (cxml:attribute "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
                   (cxml:attribute "Target" (format nil "worksheets/~A.xml" name)))))
      ;; Styles (id = count + 1)
      (let ((style-id (1+ (length sheets))))
        (cxml:with-element "Relationship"
          (cxml:attribute "Id" (format nil "rId~D" style-id))
          (cxml:attribute "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
          (cxml:attribute "Target" "styles.xml"))))))

(defun write-workbook-xml (stream &key sheets)
  "Write xl/workbook.xml to STREAM."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "workbook"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      (cxml:attribute "xmlns:r" "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
      
      (cxml:with-element "sheets"
        (loop for sh in sheets
              for i from 1
              do (let ((name (if (typep sh 'sheet) (sheet-name sh) "Sheet1"))
                       (id (if (typep sh 'sheet) (sheet-id sh) 1)))
                   (cxml:with-element "sheet"
                     (cxml:attribute "name" name)
                     (cxml:attribute "sheetId" (format nil "~D" id))
                     (cxml:attribute "r:id" (format nil "rId~D" i)))))))))

(defun write-styles-xml (stream)
  "Write minimal xl/styles.xml to STREAM."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "styleSheet"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      ;; Minimal content to satisfy Excel
      (cxml:with-element "fonts" (cxml:attribute "count" "1")
        (cxml:with-element "font"))
      (cxml:with-element "fills" (cxml:attribute "count" "2")
        (cxml:with-element "fill" (cxml:with-element "patternFill" (cxml:attribute "patternType" "none")))
        (cxml:with-element "fill" (cxml:with-element "patternFill" (cxml:attribute "patternType" "gray125"))))
      (cxml:with-element "borders" (cxml:attribute "count" "1")
        (cxml:with-element "border"))
      (cxml:with-element "cellStyleXfs" (cxml:attribute "count" "1")
         (cxml:with-element "xf" (cxml:attribute "numFmtId" "0") (cxml:attribute "fontId" "0") (cxml:attribute "fillId" "0") (cxml:attribute "borderId" "0")))
      (cxml:with-element "cellXfs" (cxml:attribute "count" "1")
         (cxml:with-element "xf" (cxml:attribute "numFmtId" "0") (cxml:attribute "fontId" "0") (cxml:attribute "fillId" "0") (cxml:attribute "borderId" "0") (cxml:attribute "xfId" "0"))))))
