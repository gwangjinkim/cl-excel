;;;; src/workbook-read.lisp

(in-package #:cl-excel)

(defclass workbook ()
  ((zip :initarg :zip :accessor workbook-zip)
   (sheets :initarg :sheets :accessor workbook-sheets)
   (shared-strings :initarg :shared-strings :accessor workbook-shared-strings)
   (styles :initarg :styles :accessor workbook-styles)
   (tables :initarg :tables :accessor workbook-tables :initform nil)
   (app-name :initarg :app-name :accessor workbook-app-name :initform nil)))

(defstruct sheet-meta
  name
  id
  rel-id)

(defun get-node-text (node)
  "Get text content of a node."
  (let ((text ""))
    (stp:do-children (child node)
      (when (typep child 'stp:text)
        (setf text (concatenate 'string text (stp:data child)))))
    text))

(defun read-app-name (zip)
  "Attempt to detect application name from docProps/app.xml or meta.xml."
  (let ((app-xml (get-entry-stream zip "docProps/app.xml"))
        (meta-xml (get-entry-stream zip "meta.xml")))
    (cond
      (app-xml 
       (let* ((dom (parse-xml app-xml))
              (app-node (find-child (stp:document-element dom) "Application")))
         (if app-node (get-node-text app-node) "")))
      (meta-xml
       (let* ((dom (parse-xml meta-xml))
              (gen-node (find-child (stp:document-element dom) "generator")))
         (if gen-node (get-node-text gen-node) "")))
      (t nil))))

(defun read-shared-strings (zip)
  "Read xl/sharedStrings.xml and return a vector of strings."
  (let ((stream (get-entry-stream zip "xl/sharedStrings.xml")))
    (if stream
        (let* ((dom (parse-xml stream))
               (items '()))
          (stp:do-children (si (stp:document-element dom))
            (when (and (typep si 'stp:element) (string= (stp:local-name si) "si"))
              ;; Logic:
              ;; 1. If <t> exists directly -> use it (standard string).
              ;; 2. If <r> (runs) exist -> iterate and concat <t> inside them.
              (let ((text ""))
                ;; Case 1: Direct <t>
                (let ((t-node (find-child si "t")))
                  (when t-node
                    (setf text (node-text t-node))))
                
                ;; Case 2: Rich text runs <r>
                (stp:do-children (child si)
                  (when (and (typep child 'stp:element) (string= (stp:local-name child) "r"))
                    (let ((rt-node (find-child child "t")))
                      (when rt-node
                        (setf text (concatenate 'string text (node-text rt-node)))))))
                
                (push text items))))
          (coerce (nreverse items) 'vector))
        #()))) ;; Return empty vector if no shared strings

(defun detect-file-format (path)
  "Detect if file is :ods, :xlsx or nil by inspecting ZIP content."
  (ignore-errors
    (zip:with-zipfile (zip path)
      (cond
        ((zip:get-zipfile-entry "mimetype" zip)
         (let ((mimetype (babel:octets-to-string 
                          (zip:zipfile-entry-contents 
                           (zip:get-zipfile-entry "mimetype" zip)))))
           (if (search "opendocument.spreadsheet" mimetype)
               :ods
               nil)))
        ((or (zip:get-zipfile-entry "xl/workbook.xml" zip)
             (zip:get-zipfile-entry "[Content_Types].xml" zip))
         :xlsx)
        (t nil)))))

(defun read-xlsx (source &key (timezone local-time:+utc-zone+))
  "Open and read an Excel workbook from SOURCE (path). Auto-detects ODS by content."
  (let ((format (detect-file-format source)))
    (case format
      (:ods (read-ods source))
      (:xlsx
       (let* ((zip (open-zip source))
              (wb (make-instance 'workbook :zip zip)))
         
         (setf (workbook-app-name wb) (read-app-name zip))
         
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
                               (let ((sh (read-sheet zip meta (workbook-shared-strings wb) (workbook-styles wb) timezone))
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
      (t (error 'xlsx-parse-error :message (format nil "Unsupported or unknown file format: ~A" source))))))

(defun close-xlsx (workbook)
  "Close the workbook and underlying resources."
  (let ((zip (workbook-zip workbook)))
    (when zip
      (close-zip zip))))

(defun open-xlsx (source &key (mode :read) (enable-cache t) (timezone local-time:+utc-zone+))
  "Open XLSX file at SOURCE.
   MODE: :read (default), :write (new file), :rw (edit existing).
   ENABLE-CACHE: Ignored for now (always true for DOM read)."
  (declare (ignore enable-cache))
  (case mode
    (:read (read-xlsx source :timezone timezone))
    (:write (make-instance 'workbook :zip nil :sheets nil))
    (:rw 
     (warn "Opening XLSX in EDIT mode (:rw). This is a best-effort implementation. 
            Charts, drawings, and complex styles may be lost upon saving.")
     (read-xlsx source :timezone timezone))
    (t (error "Unknown mode: ~A" mode))))

(defmacro with-xlsx ((var source &key (mode :read) (enable-cache t) (timezone 'local-time:+utc-zone+)) &body body)
  `(let ((,var (open-xlsx ,source :mode ,mode :enable-cache ,enable-cache :timezone ,timezone)))
     (unwind-protect
          (progn ,@body)
       (when (and ,var (not (eq ,mode :write)))
         (close-xlsx ,var)))))

(defmacro with-xlsx-save ((var source &key (mode :read) (enable-cache t) 
                                           (out-file nil)
                                           (timezone 'local-time:+utc-zone+)) 
                          &body body)
  "Like with-xlsx but also saves the workbook when done."
  `(with-xlsx (,var ,source :mode ,mode :enable-cache ,enable-cache :timezone ,timezone)
     ,@body
     (save-excel ,var ,(or out-file source))))

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

(defun sheet (workbook which)
  "Get a sheet object by name (string) or index (integer, 1-based)."
  (typecase which
    (integer 
     (if (and (>= which 1) (<= which (length (workbook-sheets workbook))))
         (nth (1- which) (workbook-sheets workbook))
         (error 'sheet-missing-error :name which)))
    (string 
     (or (find which (workbook-sheets workbook) 
               :key #'sheet-name 
               :test #'string=)
         (error 'sheet-missing-error :name which)))
    (t (error 'xlsx-error :message "Sheet selector must be string or integer."))))
