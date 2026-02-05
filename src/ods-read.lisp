(in-package #:cl-excel)

;;; ODS Reading Implementation

(defun get-ods-attribute (attrs local-name)
  "Get attribute value by local name from XMLS-style attributes list."
  (let ((entry (find local-name attrs 
                     :test (lambda (name key)
                             (if (consp key)
                                 (string= name (car key))
                                 (string= name key)))
                     :key #'car)))
    (cadr entry)))

(defun parse-ods-cell (sax)
  "Return string or number according to type of ods cell sax."
  (let ((attrs (second sax)))
    (let ((type (get-ods-attribute attrs "value-type"))
          (value-attr (get-ods-attribute attrs "value"))
          (date-attr (get-ods-attribute attrs "date-value"))
          (content (car (last (car (last sax)))))) ;; Fallback to text content
      
      (cond 
        ((string= type "float") 
         (if value-attr (parse-number:parse-number value-attr) 0.0))
        ((string= type "date")
         (if date-attr
             (local-time:parse-timestring date-attr)
             content))
        ((string= type "string")
         (or content ""))
        (t 
         (or content ""))))))

(defun number-columns-repeated (sax)
  "Check for table:number-columns-repeated attribute."
  (let ((val (get-ods-attribute (second sax) "number-columns-repeated")))
    (if val (parse-integer val) 1)))

(defun read-ods-sheet (zip content-xml sheet-name sheet-id)
  "Parse ODS content.xml for a specific sheet name."
  ;; Note: content.xml contains ALL sheets properly nested.
  ;; For efficiency we might want to parse it once and split, but for now we stream.
  
  (let ((cells (make-hash-table :test 'equal))
        (max-row 0)
        (max-col 0))
    
    (klacks:with-open-source (s (cxml:make-source content-xml))
      (loop
        (multiple-value-bind (key val) (klacks:peek s)
          (declare (ignore val))
          (cond
            ((eq key :end-document) (return))
            ((and (eq key :start-element) (string= (klacks:current-lname s) "table"))
             (let ((name (or (klacks:get-attribute s "name")
                             (klacks:get-attribute s "name" "urn:oasis:names:tc:opendocument:xmlns:table:1.0"))))
               (if (and name (string= name sheet-name))
                   ;; Found our sheet
                   (progn
                     (klacks:consume s)
                     (loop :with row-idx = 1
                           :for k = (klacks:peek s)
                           :while k
                           :do
                           (cond
                             ((and (eq k :end-element) (string= (klacks:current-lname s) "table"))
                              (return))
                             ((and (eq k :start-element) (string= (klacks:current-lname s) "table-row"))
                              (klacks:consume s)
                              (loop :with col-idx = 1
                                    :for k2 = (klacks:peek s)
                                    :while k2
                                    :do
                                    (cond
                                      ((and (eq k2 :end-element) (string= (klacks:current-lname s) "table-row"))
                                       (return))
                                      ((and (eq k2 :start-element) (string= (klacks:current-lname s) "table-cell"))
                                       (let* ((sax (klacks:serialize-element s (cxml-xmls:make-xmls-builder)))
                                              (val (parse-ods-cell sax))
                                              (repeats (number-columns-repeated sax)))
                                         
                                         (dotimes (i repeats)
                                           (when val
                                             (setf (gethash (cons row-idx col-idx) cells) 
                                                   (make-cell val :type "ods"))) ;; Basic wrapper
                                           (setf max-col (max max-col col-idx))
                                           (incf col-idx))))
                                      (t (klacks:consume s))))
                              (setf max-row (max max-row row-idx))
                              (incf row-idx))
                             (t (klacks:consume s)))))
                   (klacks:consume s)))) ;; Skip other tables
            (t (klacks:consume s))))))
    
    (make-instance 'sheet 
                   :name sheet-name 
                   :cells cells
                   :dimension (format nil "A1:~A~A" (col-name max-col) max-row))))

(defun read-ods (source)
  "Read an ODS file and return a WORKBOOK object."
  (let* ((zip (open-zip source))
         (content-xml (babel:octets-to-string 
                       (zip:zipfile-entry-contents 
                        (gethash "content.xml" (zip-entries zip))))))
    
    (let ((sheets '()))
      ;; 1. First pass: Get sheet names
      (klacks:with-open-source (s (cxml:make-source content-xml))
        (loop
          (multiple-value-bind (key val) (klacks:peek s)
            (declare (ignore val))
            (cond
              ((eq key :end-document) (return))
              ((and (eq key :start-element) (string= (klacks:current-lname s) "table"))
               (let ((name (or (klacks:get-attribute s "name") ;; try without NS
                               (klacks:get-attribute s "name" "urn:oasis:names:tc:opendocument:xmlns:table:1.0"))))
                 (when name (push name sheets)))
               (klacks:consume s))
              (t (klacks:consume s))))))
      
      (setf sheets (nreverse sheets))
      
      (let ((wb (make-instance 'workbook :zip zip)))
        (setf (workbook-sheets wb)
              (loop :for name :in sheets
                    :for i :from 1
                    :collect (read-ods-sheet zip content-xml name i)))
        wb))))
