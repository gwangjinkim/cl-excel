(in-package #:cl-excel)

(defun write-table-xml (table stream)
  "Write the XML content for a TABLE to STREAM."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "table"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      (cxml:attribute "id" (format nil "~D" (table-id table)))
      (cxml:attribute "name" (table-name table))
      (cxml:attribute "displayName" (table-display-name table))
      (cxml:attribute "ref" (table-ref table))
      (cxml:attribute "totalsRowShown" "0")
      
      (cxml:with-element "autoFilter"
        (cxml:attribute "ref" (table-ref table)))
      
      (let ((cols (table-columns table)))
        (cxml:with-element "tableColumns"
          (cxml:attribute "count" (format nil "~D" (length cols)))
          (dolist (col cols)
            (cxml:with-element "tableColumn"
              (cxml:attribute "id" (format nil "~D" (table-column-id col)))
              (cxml:attribute "name" (table-column-name col))))))
      
      (cxml:with-element "tableStyleInfo"
        (cxml:attribute "name" "TableStyleMedium9")
        (cxml:attribute "showFirstColumn" "0")
        (cxml:attribute "showLastColumn" "0")
        (cxml:attribute "showRowStripes" "1")
        (cxml:attribute "showColumnStripes" "0")))))
