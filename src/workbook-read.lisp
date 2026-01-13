;;;; src/workbook-read.lisp

(in-package #:cl-excel)

(defclass workbook ()
  ((zip :initarg :zip :accessor workbook-zip)
   (sheets :initarg :sheets :accessor workbook-sheets) ;; list of sheet structs
   (shared-strings :initarg :shared-strings :accessor workbook-shared-strings)
   (styles :initarg :styles :accessor workbook-styles)
   (tables :initarg :tables :accessor workbook-tables :initform nil)))

(defstruct sheet-meta
  name
  id
  rel-id)

(defun read-shared-strings (zip)
  "Read xl/sharedStrings.xml and return a vector of strings."
  (let ((stream (get-entry-stream zip "xl/sharedStrings.xml")))
    (if stream
        (let* ((dom (parse-xml stream))
               (items '()))
          (stp:do-children (si (stp:document-element dom))
            (when (and (typep si 'stp:element) (string= (stp:local-name si) "si"))
              ;; Basic: read <t> child. Rich text (<r>) ignored for now.
              (let ((t-node (find-child si "t")))
                (push (if t-node (node-text t-node) "") items))))
          (coerce (nreverse items) 'vector))
        #()))) ;; Return empty vector if no shared strings

(defun read-xlsx (source)
  "Open and read an Excel workbook from SOURCE (path)."
  (let* ((zip (open-zip source))
         (wb (make-instance 'workbook :zip zip)))
    
    ;; 1. Parse workbook.xml for sheets
    (let ((wb-stream (get-entry-stream zip "xl/workbook.xml")))
      (unless wb-stream
        (error 'xlsx-parse-error :message "xl/workbook.xml not found"))
      (let* ((dom (parse-xml wb-stream))
             (sheets-node (find-child (stp:document-element dom) "sheets")))
        (unless sheets-node
          (error 'xlsx-parse-error :message "No <sheets> element in workbook.xml"))
        
        ;; First pass: Metadata only
        (let ((metas (mapcar (lambda (node)
                               (make-sheet-meta 
                                :name (get-attribute node "name")
                                :id (get-attribute node "sheetId")
                                :rel-id (get-attribute node "id"))) ;; r:id
                             (find-children sheets-node "sheet"))))
          
          ;; 2. Parse shared strings
          (setf (workbook-shared-strings wb) (read-shared-strings zip))
          
          ;; 3. Parse styles (M5)
          (setf (workbook-styles wb) (read-styles zip))
          
          ;; 4. Load actual sheets
          (setf (workbook-sheets wb)
                (mapcar (lambda (meta)
                          (let ((sh (read-sheet zip meta (workbook-shared-strings wb) (workbook-styles wb)))
                                (rels-path (format nil "xl/worksheets/_rels/sheet~A.xml.rels" (sheet-meta-id meta))))
                            
                            ;; M5: Load Table Relationships
                            (let ((rels (read-relationships zip rels-path)))
                              (when rels 
                                (let ((tables '()))
                                  (maphash (lambda (rid target)
                                             (declare (ignore rid))
                                             ;; Target resolution
                                             ;; Simple normalization: if it contains "table", try to find entry.
                                             (let ((path (if (uiop:string-prefix-p "../" target)
                                                             (concatenate 'string "xl/" (subseq target 3))
                                                             (if (uiop:string-prefix-p "/xl/" target)
                                                                 (subseq target 1)
                                                                 (concatenate 'string "xl/worksheets/" target)))))
                                               
                                               (when (or (search "table" path) (search "Table" path))
                                                 (let ((stream (get-entry-stream zip path)))
                                                   (when stream
                                                     (let ((tbl (read-table-xml stream :sheet sh)))
                                                       (push tbl tables)))))))
                                           rels)
                                  (setf (sheet-tables sh) tables))))
                            sh))
                        metas))
          
          ;; 5. Collect all tables into workbook-tables (scan all sheets)
          (let ((all-tables '()))
            (dolist (sh (workbook-sheets wb))
              (dolist (tbl (sheet-tables sh))
                (push tbl all-tables)))
            (setf (workbook-tables wb) (nreverse all-tables))))))
    wb))

(defun close-xlsx (workbook)
  "Close the workbook and underlying resources."
  (close-zip (workbook-zip workbook)))

(defun open-xlsx (source &key (mode :read) (enable-cache t))
  "Open XLSX file at SOURCE.
   MODE: :read (default), :write (new file), :rw (edit existing).
   ENABLE-CACHE: Ignored for now (always true for DOM read)."
  (declare (ignore enable-cache))
  (case mode
    (:read (read-xlsx source))
    (:write (make-instance 'workbook :zip nil :sheets nil))
    (:rw 
     (warn "Opening XLSX in EDIT mode (:rw). This is a best-effort implementation. 
            Charts, drawings, and complex styles may be lost upon saving.")
     (read-xlsx source))
    (t (error "Unknown mode: ~A" mode))))

(defmacro with-xlsx ((var source &key (mode :read) (enable-cache t)) &body body)
  `(let ((,var (open-xlsx ,source :mode ,mode :enable-cache ,enable-cache)))
     (unwind-protect
          (progn ,@body)
       (when (and (workbook-zip ,var) (not (eq ,mode :write)))
         (close-xlsx ,var)))))

;;; Public API impl

(defun sheet-names (workbook)
  (mapcar #'sheet-name (workbook-sheets workbook)))

(defun sheet-count (workbook)
  (length (workbook-sheets workbook)))

(defun has-sheet-p (workbook name)
  (and (find name (workbook-sheets workbook) 
             :key #'sheet-name 
             :test #'string=)
       t))
