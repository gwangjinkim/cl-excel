(require :asdf)
(asdf:load-system :cl-excel)

(defun test-smart ()
  (format t "Testing Smart API (M11)...~%")
  
  ;; Setup: Create a file with some data
  (let ((wb (make-instance 'cl-excel::workbook))
        (sheet (make-instance 'cl-excel:sheet :name "SmartSheet" :id 1)))
    (setf (cl-excel::workbook-sheets wb) (list sheet))
    
    ;; A1..A3
    (setf (cl-excel:val sheet "A1") 10)
    (setf (cl-excel:val sheet "A2") 20)
    (setf (cl-excel:val sheet "A3") 30)
    ;; B2
    (setf (cl-excel:val sheet "B2") "B-Val")
    
    (cl-excel:save-excel wb "smart.xlsx")
    (format t "Saved smart.xlsx~%"))
    
  ;; Test 1: list-sheets
  (let ((sheets (cl-excel:list-sheets "smart.xlsx")))
    (format t "Sheets: ~A~%" sheets)
    (unless (equal sheets '("SmartSheet"))
      (error "list-sheets failed")))
      
  ;; Test 2: Smart Ranges
  
  ;; "A" -> Column 1 trimmed (A1:A3, ignore A4.. but technically used-range of sheet includes B2 so max row is 3. 
  ;; Wait, used-range bounding box is A1:B3. 
  ;; So Column A trimmed should be A1:A3.
  (let ((col-a (cl-excel:read-file "smart.xlsx" 1 "A")))
    (format t "Col A: ~A~%" col-a)
    ;; Should be ((10) (20) (30))
    ;; Note: `used-range` calculates global bbox. 
    ;; A1..A3 are present. B2 is present. Max row is 3. 
    ;; So range is A1:A3.
    (unless (equal col-a '((10) (20) (30)))
      (error "Smart range 'A' failed")))

  ;; 1 -> Column 1 (same as A)
  (let ((col-1 (cl-excel:read-file "smart.xlsx" 1 1)))
    (format t "Col 1: ~A~%" col-1)
    (unless (equal col-1 '((10) (20) (30)))
      (error "Smart range 1 failed")))
      
  ;; "A1" -> Single cell ((10))
  (let ((cell (cl-excel:read-file "smart.xlsx" 1 "A1")))
    (format t "Cell A1: ~A~%" cell)
    (unless (equal cell '((10)))
      (error "Smart range 'A1' failed")))
  
  (format t "SUCCESS.~%"))

(test-smart)
(sb-ext:exit)
