(in-package #:cl-excel)

;;; ODS Writing Implementation

(defun content-xml-header ()
  "<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<office:document-content xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" xmlns:table=\"urn:oasis:names:tc:opendocument:xmlns:table:1.0\" xmlns:text=\"urn:oasis:names:tc:opendocument:xmlns:text:1.0\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:number=\"urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0\" xmlns:style=\"urn:oasis:names:tc:opendocument:xmlns:style:1.0\" office:version=\"1.2\">
  <office:body>
    <office:spreadsheet>")

(defun content-xml-footer ()
  "    </office:spreadsheet>
  </office:body>
</office:document-content>")

(defun manifest-xml ()
  "<?xml version=\"1.0\" encoding=\"UTF-8\"?>
<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\" manifest:version=\"1.2\">
 <manifest:file-entry manifest:full-path=\"/\" manifest:version=\"1.2\" manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/>
 <manifest:file-entry manifest:full-path=\"content.xml\" manifest:media-type=\"text/xml\"/>
 <manifest:file-entry manifest:full-path=\"styles.xml\" manifest:media-type=\"text/xml\"/>
 <manifest:file-entry manifest:full-path=\"meta.xml\" manifest:media-type=\"text/xml\"/>
</manifest:manifest>")

(defun write-ods-cell (sink val)
  "Write an ODS cell using CXML sink."
  (cond
    ((numberp val)
     (cxml:with-element "table:table-cell"
       (cxml:attribute "office:value-type" "float")
       (cxml:attribute "office:value" (format nil "~F" val))
       (cxml:with-element "text:p" (cxml:text (format nil "~F" val)))))
    ((typep val 'local-time:timestamp)
     (let ((iso (local-time:format-timestring nil val :format local-time:+iso-8601-format+)))
       (cxml:with-element "table:table-cell"
         (cxml:attribute "office:value-type" "date")
         (cxml:attribute "office:date-value" iso)
         (cxml:with-element "text:p" (cxml:text iso)))))
    (t
     (cxml:with-element "table:table-cell"
       (cxml:attribute "office:value-type" "string")
       (cxml:with-element "text:p" (cxml:text (princ-to-string val)))))))

(defun write-ods-sheet (sink sheet-name data)
  "Write an ODS sheet using CXML sink."
  (cxml:with-element "table:table"
    (cxml:attribute "table:name" sheet-name)
    (dolist (row data)
      (cxml:with-element "table:table-row"
        (dolist (cell row)
          (write-ods-cell sink cell))))))

(defun write-ods (data-alist filepath)
  "Write internal data representation (alist of \"SheetName\" -> rows) to ODS."
  (zip:with-output-to-zipfile (zip filepath)
    
    ;; 1. mimetype (first, uncompressed, MUST be first)
    (let ((bytes (babel:string-to-octets "application/vnd.oasis.opendocument.spreadsheet")))
      (flexi-streams:with-input-from-sequence (s bytes)
        (zip:write-zipentry zip "mimetype" s :file-write-date (get-universal-time))))
    
    ;; 2. content.xml
    (let* ((bytes (flexi-streams:with-output-to-sequence (out-raw)
                    (let ((out (flexi-streams:make-flexi-stream out-raw :external-format :utf-8)))
                      (cxml:with-xml-output (cxml:make-character-stream-sink out :canonical nil)
                        (cxml:with-element "office:document-content"
                          (cxml:attribute "xmlns:office" "urn:oasis:names:tc:opendocument:xmlns:office:1.0")
                          (cxml:attribute "xmlns:table" "urn:oasis:names:tc:opendocument:xmlns:table:1.0")
                          (cxml:attribute "xmlns:text" "urn:oasis:names:tc:opendocument:xmlns:text:1.0")
                          (cxml:attribute "xmlns:xsd" "http://www.w3.org/2001/XMLSchema")
                          (cxml:attribute "xmlns:number" "urn:oasis:names:tc:opendocument:xmlns:datastyle:1.0")
                          (cxml:attribute "xmlns:style" "urn:oasis:names:tc:opendocument:xmlns:style:1.0")
                          (cxml:attribute "office:version" "1.2")
                          (cxml:with-element "office:body"
                            (cxml:with-element "office:spreadsheet"
                              (dolist (sheet-entry data-alist)
                                (write-ods-sheet nil (car sheet-entry) (cdr sheet-entry)))))))))))
      (flexi-streams:with-input-from-sequence (s bytes)
        (zip:write-zipentry zip "content.xml" s :file-write-date (get-universal-time))))
      
    ;; 3. META-INF/manifest.xml
    (let ((bytes (babel:string-to-octets (manifest-xml))))
      (flexi-streams:with-input-from-sequence (s bytes)
        (zip:write-zipentry zip "META-INF/manifest.xml" s :file-write-date (get-universal-time))))

    ;; 4. Dummy styles and meta (also using cxml for safety)
    (let ((bytes (babel:string-to-octets "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-styles xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" office:version=\"1.2\"/>")))
      (flexi-streams:with-input-from-sequence (s bytes)
        (zip:write-zipentry zip "styles.xml" s :file-write-date (get-universal-time))))
      
    (let ((bytes (babel:string-to-octets "<?xml version=\"1.0\" encoding=\"UTF-8\"?><office:document-meta xmlns:office=\"urn:oasis:names:tc:opendocument:xmlns:office:1.0\" office:version=\"1.2\"/>")))
      (flexi-streams:with-input-from-sequence (s bytes)
        (zip:write-zipentry zip "meta.xml" s :file-write-date (get-universal-time))))))
