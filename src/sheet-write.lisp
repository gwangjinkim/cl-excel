;;;; src/sheet-write.lisp

(in-package #:cl-excel)

(defun write-sheet-xml (sheet stream)
  "Write the XML content for a SHEET to STREAM."
  (cxml:with-xml-output (cxml:make-character-stream-sink stream :canonical nil)
    (cxml:with-element "worksheet"
      (cxml:attribute "xmlns" "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
      (cxml:attribute "xmlns:r" "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
      
      ;; Sheet Data
      (cxml:with-element "sheetData"
        ;; To write rows efficiently, we need to sort or iterate the hash table.
        ;; Since cells are (cons row col), we need to group by row first.
        ;; For M6 large output might be slow, but correctness first.
        (let ((rows (make-hash-table :test 'eql)))
          ;; Group cells by row index
          (maphash (lambda (coord cell)
                     (let ((r (car coord))
                           (c (cdr coord)))
                       (push (cons c cell) (gethash r rows))))
                   (sheet-cells sheet))
          
          ;; Iterate sorted rows
          (let ((row-indices (sort (alexandria:hash-table-keys rows) #'<)))
            (dolist (r row-indices)
              (cxml:with-element "row"
                (cxml:attribute "r" (format nil "~D" r))
                
                ;; Sort columns in this row
                (let ((col-cells (sort (gethash r rows) #'< :key #'car)))
                  (dolist (pair col-cells)
                    (let* ((c (car pair))
                           (cell (cdr pair))
                           (val (cell-value cell))
                           (ref (format nil "~A~D" (col-name c) r))
                           (type (cell-type cell)))
                      
                      (cxml:with-element "c"
                        (cxml:attribute "r" ref)
                        
                        ;; Determine type if not set (simple inference for new cells)
                        (let ((final-type (or type 
                                             (cond 
                                               ((numberp val) "n")
                                               ((stringp val) "inlineStr") 
                                               ((typep val 'boolean) "b")
                                               (t "inlineStr")))))
                          
                          (cxml:attribute "t" final-type)
                          
                          (cond
                            ((string= final-type "inlineStr")
                             (cxml:with-element "is"
                               (cxml:with-element "t" (cxml:text (format nil "~A" val)))))
                            ((string= final-type "b")
                             (cxml:with-element "v" (cxml:text (if val "1" "0"))))
                            (t ;; Number / default
                             (cxml:with-element "v" (cxml:text (format nil "~A" val)))))))))))))))))) ; Close sheetData, worksheet
