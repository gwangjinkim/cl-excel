;;;; src/rels.lisp

(in-package #:cl-excel)

(defun read-relationships (zip path)
  "Read a .rels file from ZIP at PATH. 
   Returns a hash-table mapping Relationship ID (r:id) to Target path."
  ;; .rels paths are usually like "xl/worksheets/_rels/sheet1.xml.rels"
  (let ((stream (get-entry-stream zip path))
        (rels (make-hash-table :test 'equal)))
    (when stream
      (let* ((dom (parse-xml stream))
             (root (stp:document-element dom)))
        (stp:do-children (child root)
          (when (and (typep child 'stp:element)
                     (string= (stp:local-name child) "Relationship"))
            (let ((id (get-attribute child "Id"))
                  (target (get-attribute child "Target"))
                  (type (get-attribute child "Type")))
              ;; We mostly care about ID -> Target.
              ;; Path normalization might be needed if Target is relative.
              ;; For standard Excel, targets in sheet rels are usually relative to xl/worksheets/
              ;; or absolute like "/xl/tables/table1.xml".
              ;; We will store the raw target for now and handle resolution in caller.
              (setf (gethash id rels) target))))))
    rels))
