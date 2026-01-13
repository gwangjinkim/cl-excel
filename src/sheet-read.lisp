;;;; src/sheet-read.lisp

(in-package #:cl-excel)


(defun excel-date-to-timestamp (serial)
  "Convert Excel serial date (float/int) to local-time:timestamp.
   Assumes 1900 date system."
  (let* ((days (floor serial))
         (frac (- serial days))
         ;; Excel 1900 epoch is Dec 30 1899.
         ;; But there's the 1900 leap year bug (day 60).
         ;; For modern dates (>60), effectively Dec 30 1899.
         (base (local-time:encode-timestamp 0 0 0 0 30 12 1899 :timezone local-time:+utc-zone+)))
         
    (local-time:timestamp+ base 
                           (+ (* days 86400) 
                              (round (* frac 86400))) 
                           :sec)))

(defun parse-cell-value (type raw-val shared-strings style-id styles)
  "Parse raw value string based on cell type and style."
  (cond
    ((string= type "s") ;; Shared string
     (let ((idx (parse-integer raw-val)))
       (if (and shared-strings (< idx (length shared-strings)))
           (aref shared-strings idx)
           raw-val)))
    ((string= type "n") ;; Number
     (let ((num (cond 
                  ((find #\. raw-val) (parse-number:parse-number raw-val))
                  (t (parse-integer raw-val)))))
       ;; Check for Date style
       (if (and styles style-id)
           (let* ((xf-id style-id)
                  (num-fmt-id (get-style-num-fmt styles xf-id)))
             (if (and num-fmt-id 
                      (or (<= 14 num-fmt-id 22) ;; Built-in date formats
                          (= num-fmt-id 45) (= num-fmt-id 46) (= num-fmt-id 47)
                          (let ((fmt-code (gethash num-fmt-id (styles-num-fmts styles))))
                            (is-date-format fmt-code))))
                 (excel-date-to-timestamp num)
                 num))
           num)))
    ((string= type "b") ;; Boolean
     (string= raw-val "1"))
    ((string= type "str") raw-val) ;; Formula string
    ((string= type "inlineStr") raw-val)
    ((string= type "e") raw-val) ;; Error
    (t ;; Default (no type attribute usually means number)
     (if (and raw-val (> (length raw-val) 0))
         ;; Recurse as number type
         (parse-cell-value "n" raw-val shared-strings style-id styles)
         nil))))

(defun read-sheet (zip sheet-meta shared-strings styles)
  "Read and parse a worksheet from the ZIP file."
  (let ((path (format nil "xl/worksheets/sheet~A.xml" (sheet-meta-id sheet-meta))))
    (let ((stream (get-entry-stream zip path)))
      (unless stream
        (error 'sheet-missing-error :name (sheet-meta-name sheet-meta)))
      
      (let* ((dom (parse-xml stream))
             (root (stp:document-element dom))
             (sheet-data (find-child root "sheetData"))
             (cells (make-hash-table :test 'equal))
             (sheet-obj (make-instance 'sheet 
                                       :name (sheet-meta-name sheet-meta)
                                       :id (sheet-meta-id sheet-meta)
                                       :rel-id (sheet-meta-rel-id sheet-meta)
                                       :cells cells)))
        
        ;; Parse Dimension (optional but good)
        (let ((dim-node (find-child root "dimension")))
          (when dim-node
            (setf (sheet-dimension sheet-obj) 
                  (get-attribute dim-node "ref"))))
        
        (when sheet-data
          (stp:do-children (row-node sheet-data)
            (when (and (typep row-node 'stp:element) (string= (stp:local-name row-node) "row"))
              (stp:do-children (c-node row-node)
                (when (and (typep c-node 'stp:element) (string= (stp:local-name c-node) "c"))
                  (let* ((ref (get-attribute c-node "r"))
                         (type (or (get-attribute c-node "t") "n")) ;; Default to number if 't' missing
                         (style-id-attr (get-attribute c-node "s"))
                         (style-id (if style-id-attr (parse-integer style-id-attr) nil))
                         (v-node (find-child c-node "v"))
                         ;; Inline string support 
                         (is-node (find-child c-node "is"))
                         (raw-val (cond 
                                    (v-node (node-text v-node))
                                    (is-node 
                                     (let ((t-node (find-child is-node "t")))
                                       (if t-node (node-text t-node) "")))
                                    (t nil))))
                    
                    (when raw-val
                      (let ((val (parse-cell-value type raw-val shared-strings style-id styles))
                            (parsed-ref (parse-cell-ref ref)))
                        (setf (gethash (cons (cell-ref-row parsed-ref) 
                                             (cell-ref-col parsed-ref)) 
                                       cells)
                              (make-cell val :type type :style-id style-id))))))))))
        
        
        sheet-obj))))

;;; Lazy Reading (Klacks)

(defun make-sheet-iterator (zip sheet-meta shared-strings styles)
  "Create a lazy iterator for rows. Returns a closure that returns the next 
   list of cell values (or NIL if done), and a cleanup thunk."
  (let* ((path (format nil "xl/worksheets/sheet~A.xml" (sheet-meta-id sheet-meta)))
         (stream (get-entry-stream zip path)))
    (unless stream
      (error 'sheet-missing-error :name (sheet-meta-name sheet-meta)))
    
    (let ((source (cxml:make-source stream)))
      (values 
       (lambda () 
         (block :next-row
           (loop
             (multiple-value-bind (key val) (klacks:peek source)
               (declare (ignore val))
               (cond
                 ((eq key :end-document) (return-from :next-row nil))
                 ((and (eq key :start-element) (string= (klacks:current-lname source) "row"))
                  ;; In a row
                  (klacks:consume source) ;; consume <row>
                  (let ((row-vals '()))
                    (loop
                      (multiple-value-bind (k v) (klacks:peek source)
                        (declare (ignore v))
                        (cond
                          ((and (eq k :end-element) (string= (klacks:current-lname source) "row"))
                           (klacks:consume source) ;; consume </row>
                           (return)) ;; Done with row
                          ((and (eq k :start-element) (string= (klacks:current-lname source) "c"))
                           ;; Found cell
                           (let* ((ref (klacks:get-attribute source "r"))
                                  (type (or (klacks:get-attribute source "t") "n"))
                                  (s-attr (klacks:get-attribute source "s"))
                                  (style-id (if s-attr (parse-integer s-attr) nil))
                                  (raw-val nil))
                             (declare (ignore ref))
                             (klacks:consume source) ;; consume <c>
                             
                             ;; Parse children of <c> (v, is, etc)
                             (loop
                               (multiple-value-bind (ck cv) (klacks:peek source)
                                 (declare (ignore cv))
                                 (cond
                                   ((and (eq ck :end-element) (string= (klacks:current-lname source) "c"))
                                    (klacks:consume source)
                                    (return))
                                   ((and (eq ck :start-element) (string= (klacks:current-lname source) "v"))
                                    (klacks:consume source)
                                    ;; Read text content
                                    (let ((txt ""))
                                      (loop
                                        (multiple-value-bind (tk tv) (klacks:peek source)
                                          (if (eq tk :characters)
                                              (progn (setf txt (concatenate 'string txt tv))
                                                     (klacks:consume source))
                                              (return))))
                                      (setf raw-val txt))
                                    ;; Consume </v>
                                    (klacks:find-event source :end-element))
                                   ((and (eq ck :start-element) (string= (klacks:current-lname source) "is"))
                                    (klacks:consume source)
                                    (loop
                                         (multiple-value-bind (isk isv) (klacks:peek source)
                                           (declare (ignore isv))
                                           (cond 
                                             ((and (eq isk :end-element) (string= (klacks:current-lname source) "is"))
                                              (klacks:consume source)
                                              (return))
                                             ((and (eq isk :start-element) (string= (klacks:current-lname source) "t"))
                                              (klacks:consume source)
                                              (let ((txt ""))
                                                (loop (multiple-value-bind (tk tv) (klacks:peek source)
                                                        (if (eq tk :characters)
                                                            (progn (setf txt (concatenate 'string txt tv))
                                                                   (klacks:consume source))
                                                            (return))))
                                                (setf raw-val txt))
                                               (klacks:find-event source :end-element))
                                             (t (klacks:consume source))))))
                                   (t (klacks:consume source)))))
                             
                             ;; Parse value and push
                             (when raw-val
                               (let ((val (parse-cell-value type raw-val shared-strings style-id styles)))
                                 (push val row-vals))))
                          (t (klacks:consume source)))))
                    (return-from :next-row (nreverse row-vals))))
                 (t (klacks:consume source)))))))
       (lambda () (close stream))))))
