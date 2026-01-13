(in-package #:cl-excel)

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
