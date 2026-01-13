(require :asdf)
(asdf:load-system :cl-excel)
(asdf:load-system :zip)

(defun test-edit-mode ()
  (format t "Testing Edit Mode (:rw)...~%")
  
  ;; 1. Create Original File
  (let ((wb (make-instance 'cl-excel::workbook))
        (sheet (make-instance 'cl-excel::sheet :name "OriginalSheet" :id 1)))
    (setf (cl-excel::workbook-sheets wb) (list sheet))
    (setf (cl-excel:get-cell sheet "A1") (cl-excel::make-cell "Original Value"))
    (setf (cl-excel:get-cell sheet "B1") (cl-excel::make-cell 42))
    
    (cl-excel:write-xlsx wb "original.xlsx")
    (format t "Created original.xlsx~%"))
  
  ;; 2. Open in RW mode and Modify
  (cl-excel:with-xlsx (wb "original.xlsx" :mode :rw)
    (format t "Opened original.xlsx in RW mode.~%")
    (let ((sh (cl-excel:sheet wb 1)))
      (format t "Cell A1: ~A~%" (cl-excel:cell-value (cl-excel:get-cell sh "A1")))
      
      ;; Modify
      (setf (cl-excel:get-cell sh "A1") (cl-excel::make-cell "Modified Value"))
      (setf (cl-excel:get-cell sh "C1") (cl-excel::make-cell "New Cell"))
      
      ;; Save to new file
      (cl-excel:write-xlsx wb "edited.xlsx")
      (format t "Saved to edited.xlsx~%")))
  
  ;; 3. Verify Changes
  (cl-excel:with-xlsx (wb "edited.xlsx") ;; Read mode default
    (let ((sh (cl-excel:sheet wb 1)))
      (let ((val-a1 (cl-excel:cell-value (cl-excel:get-cell sh "A1")))
            (val-c1 (cl-excel:cell-value (cl-excel:get-cell sh "C1"))))
        
        (format t "Read check A1: ~A~%" val-a1)
        (format t "Read check C1: ~A~%" val-c1)
        
        (unless (string= val-a1 "Modified Value")
          (error "Verification Failed: A1 != Modified Value"))
        (unless (string= val-c1 "New Cell")
          (error "Verification Failed: C1 != New Cell")))))
  
  (format t "SUCCESS.~%"))

(test-edit-mode)
(sb-ext:exit)
