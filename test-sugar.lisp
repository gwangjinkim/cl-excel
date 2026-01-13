(require :asdf)
(asdf:load-system :cl-excel)

(defun test-sugar ()
  (format t "Testing Sugar API (M10)...~%")
  
  ;; 1. Write File (save-excel)
  (let ((wb (make-instance 'cl-excel::workbook))
        (sheet (make-instance 'cl-excel:sheet :name "SugarSheet" :id 1)))
    (setf (cl-excel::workbook-sheets wb) (list sheet))
    
    (setf (cl-excel:val sheet "A1") "Name")
    (setf (cl-excel:val sheet "B1") "Value")
    (setf (cl-excel:val sheet "A2") "ItemX")
    (setf (cl-excel:val sheet "B2") 100)
    (setf (cl-excel:[] sheet "A3") "ItemY") ;; Test [] alias
    (setf (cl-excel:[] sheet "B3") 200)
    
    (cl-excel:save-excel wb "sugar.xlsx")
    (format t "Saved sugar.xlsx~%"))
  
  ;; 2. Read File (read-excel, read-file)
  (let ((wb (cl-excel:read-excel "sugar.xlsx"))
        (sheet (cl-excel:sheet-of (cl-excel:read-excel "sugar.xlsx") "SugarSheet")))
    
    (let ((val-a2 (cl-excel:val sheet "A2"))
          (val-b3 (cl-excel:[] sheet "B3")))
      (format t "A2: ~A~%" val-a2)
      (format t "B3: ~A~%" val-b3)
      
      (unless (string= val-a2 "ItemX") (error "Sugar val failed A2"))
      (unless (= val-b3 200) (error "Sugar [] failed B3"))))
  
  ;; 3. read-file quick access
  (let ((rows (cl-excel:read-file "sugar.xlsx")))
    (format t "read-file all: ~A~%" rows)
    (unless (= (length rows) 3) (error "read-file length mismatch")))

  (let ((partial (cl-excel:read-file "sugar.xlsx" 1 "A2:B2")))
    (format t "read-file range A2:B2: ~A~%" partial)
    (unless (and (= (length partial) 1) 
                 (equal (first partial) '("ItemX" 100)))
      (error "read-file range mismatch")))
      
  ;; 4. map-rows
  (cl-excel:with-xlsx (wb "sugar.xlsx")
    (let ((s (cl-excel:sheet-of wb 1)))
       (let ((mapped (cl-excel:map-rows (lambda (r) (first r)) s)))
         (format t "mapped first col: ~A~%" mapped)
         (unless (equal mapped '("Name" "ItemX" "ItemY"))
           (error "map-rows failed")))))

  (format t "SUCCESS.~%"))

(test-sugar)
(sb-ext:exit)
