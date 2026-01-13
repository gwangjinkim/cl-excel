;;;; src/workbook-write.lisp

(in-package #:cl-excel)

(defun write-content-types-xml (stream)
  "Write [Content_Types].xml to STREAM."
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
      
      ;; Add sheets ?? Since names are dynamic, we usually iterate logic.
      ;; For now, assume a fixed structure or pass sheet list?
      ;; To keep it decoupled, we might need a better architecture, but let's assume
      ;; default sheet1.
      (cxml:with-element "Override"
        (cxml:attribute "PartName" "/xl/worksheets/sheet1.xml")
        (cxml:attribute "ContentType" "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml")))))

(defun write-rels-xml (stream)
  "Write _rels/.rels to STREAM."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "Relationships"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/package/2006/relationships")
      (cxml:with-element "Relationship"
        (cxml:attribute "Id" "rId1")
        (cxml:attribute "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument")
        (cxml:attribute "Target" "xl/workbook.xml")))))

(defun write-workbook-rels-xml (stream)
  "Write xl/_rels/workbook.xml.rels to STREAM."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "Relationships"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/package/2006/relationships")
      (cxml:with-element "Relationship"
        (cxml:attribute "Id" "rId1")
        (cxml:attribute "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet")
        (cxml:attribute "Target" "worksheets/sheet1.xml"))
      (cxml:with-element "Relationship"
        (cxml:attribute "Id" "rId2")
        (cxml:attribute "Type" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles")
        (cxml:attribute "Target" "styles.xml")))))

(defun write-workbook-xml (stream)
  "Write xl/workbook.xml to STREAM."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "workbook"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      (cxml:attribute "xmlns:r" "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
      
      (cxml:with-element "sheets"
        (cxml:with-element "sheet"
          (cxml:attribute "name" "Sheet1")
          (cxml:attribute "sheetId" "1")
          (cxml:attribute "r:id" "rId1"))))))

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
