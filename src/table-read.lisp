;;;; src/table-read.lisp

(in-package #:cl-excel)

(defun read-table-xml (stream &key sheet)
  "Parse a table definition XML stream.
   SHEET is the parent sheet object (optional, but needed for back-reference)."
  (let* ((dom (parse-xml stream))
         (root (stp:document-element dom))
         (cols-node (find-child root "tableColumns"))
         (table (make-instance 'table
                               :id (parse-integer (get-attribute root "id"))
                               :name (get-attribute root "name")
                               :display-name (get-attribute root "displayName")
                               :ref (get-attribute root "ref")
                               :sheet sheet)))
    (when cols-node
      (let ((cols '()))
        (stp:do-children (c-node cols-node)
          (when (and (typep c-node 'stp:element) (string= (stp:local-name c-node) "tableColumn"))
            (push (make-table-column 
                   :id (parse-integer (get-attribute c-node "id"))
                   :name (get-attribute c-node "name"))
                  cols)))
        (setf (table-columns table) (nreverse cols))))
    table))
