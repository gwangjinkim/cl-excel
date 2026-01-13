;;;; src/sheet-read.lisp
(in-package #:cl-excel)

(defun excel-date-to-timestamp (serial)
  "Convert Excel serial date (float/int) to local-time:timestamp.
   Assumes 1900 date system."
  (let* ((days (floor serial))
         (frac (- serial days))
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
       (if (and styles style-id)
           (let* ((xf-id style-id)
                  (num-fmt-id (get-style-num-fmt styles xf-id)))
             (if (and num-fmt-id 
                      (or (<= 14 num-fmt-id 22)
                          (= num-fmt-id 45) (= num-fmt-id 46) (= num-fmt-id 47)
                          (let ((fmt-code (gethash num-fmt-id (styles-num-fmts styles))))
                            (is-date-format fmt-code))))
                 (excel-date-to-timestamp num)
                 num))
           num)))
    ((string= type "b") (string= raw-val "1"))
    ((string= type "str") raw-val)
    ((string= type "inlineStr") raw-val)
    ((string= type "e") raw-val)
    (t 
     (if (and raw-val (> (length raw-val) 0))
         (parse-cell-value "n" raw-val shared-strings style-id styles)
         nil))))

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
                  (klacks:consume source)
                  (let ((row-vals '()))
                    (loop
                      (multiple-value-bind (k v) (klacks:peek source)
                        (declare (ignore v))
                        (cond
                          ((and (eq k :end-element) (string= (klacks:current-lname source) "row"))
                           (klacks:consume source)
                           (return))
                          ((and (eq k :start-element) (string= (klacks:current-lname source) "c"))
                           (let* ((ref (klacks:get-attribute source "r"))
                                  (type (or (klacks:get-attribute source "t") "n"))
                                  (s-attr (klacks:get-attribute source "s"))
                                  (style-id (if s-attr (parse-integer s-attr) nil))
                                  (raw-val nil))
                             (declare (ignore ref))
                             (klacks:consume source)
                             (loop
                               (multiple-value-bind (ck cv) (klacks:peek source)
                                 (declare (ignore cv))
                                 (cond
                                   ((and (eq ck :end-element) (string= (klacks:current-lname source) "c"))
                                    (klacks:consume source)
                                    (return))
                                   ((and (eq ck :start-element) (string= (klacks:current-lname source) "v"))
                                    (klacks:consume source)
                                    (let ((txt ""))
                                      (loop
                                        (multiple-value-bind (tk tv) (klacks:peek source)
                                          (if (eq tk :characters)
                                              (progn (setf txt (concatenate 'string txt tv))
                                                     (klacks:consume source))
                                              (return))))
                                      ;; Consume </v>
                                    (klacks:find-event source :end-element)))
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
                             (when raw-val
                               (let ((val (parse-cell-value type raw-val shared-strings style-id styles)))
                                 (push val row-vals)))))
                          (t (klacks:consume source)))))
                    (return-from :next-row (nreverse row-vals))))
                 (t (klacks:consume source)))))))
       (lambda () (close stream))))))

(defun find-child-local (name node)
  (cxml-stp:do-children (child node)
    (when (and (typep child 'cxml-stp:element)
               (string= (cxml-stp:local-name child) name))
      (return child))))

(defun read-sheet (zip sheet-meta shared-strings styles)
  "Read and parse a worksheet from the ZIP file."
  (let ((path (format nil "xl/worksheets/sheet~A.xml" (sheet-meta-id sheet-meta))))
    (let ((stream (get-entry-stream zip path)))
      (unless stream
        (error 'sheet-missing-error :name (sheet-meta-name sheet-meta)))
      
      (let* ((dom (cxml:parse-stream stream (cxml-stp:make-builder)))
             (root (cxml-stp:document-element dom))
             (sheet-data (find-child-local "sheetData" root))
             (cells (make-hash-table :test 'equal))
             (sheet-obj (make-instance 'sheet 
                                       :name (sheet-meta-name sheet-meta)
                                       :id (sheet-meta-id sheet-meta)
                                       :rel-id (sheet-meta-rel-id sheet-meta)
                                       :cells cells)))
        
        (let ((dim-node (find-child-local "dimension" root)))
          (when dim-node
            (setf (sheet-dimension sheet-obj) 
                  (cxml-stp:attribute-value dim-node "ref"))))
        
        (when sheet-data
          (cxml-stp:do-children (row-node sheet-data)
            (when (and (typep row-node 'cxml-stp:element) (string= (cxml-stp:local-name row-node) "row"))
              (cxml-stp:do-children (c-node row-node)
                (when (and (typep c-node 'cxml-stp:element) (string= (cxml-stp:local-name c-node) "c"))
                  (let* ((ref (cxml-stp:attribute-value c-node "r"))
                         (type (or (cxml-stp:attribute-value c-node "t") "n")) 
                         (style-id-attr (cxml-stp:attribute-value c-node "s"))
                         (style-id (if style-id-attr (parse-integer style-id-attr) nil))
                         (v-node (find-child-local "v" c-node))
                         (is-node (find-child-local "is" c-node))
                         (raw-val (cond 
                                    (v-node (cxml-stp:data (cxml-stp:first-child v-node)))
                                    (is-node 
                                     (let ((t-node (find-child-local "t" is-node)))
                                       (if t-node (cxml-stp:data (cxml-stp:first-child t-node)) "")))
                                    (t nil))))
                    
                    (when raw-val
                      (let ((val (parse-cell-value type raw-val shared-strings style-id styles))
                            (parsed-ref (parse-cell-ref ref)))
                        (setf (gethash (cons (cell-ref-row parsed-ref) 
                                             (cell-ref-col parsed-ref)) 
                                       cells)
                              (make-cell val :type type :style-id style-id))))))))))
        
        sheet-obj))))
