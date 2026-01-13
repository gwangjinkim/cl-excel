;;;; src/xml.lisp

(in-package #:cl-excel)

(defun parse-xml (stream)
  "Parse XML from STREAM into a CXML-STP DOM document."
  (cxml:parse-stream stream (stp:make-builder)))

(defun find-child (node name)
  "Find the first child element of NODE with tag name NAME (string)."
  (stp:find-child-if (lambda (child)
                       (and (typep child 'stp:element)
                            (string= (stp:local-name child) name)))
                     node))

(defun find-children (node name)
  "Return a list of all child elements of NODE with tag name NAME."
  (stp:filter-children (lambda (n)
                         (and (typep n 'stp:element)
                              (string= (stp:local-name n) name)))
                       node))

(defun get-attribute (node name)
  "Get attribute value by NAME from NODE."
  (stp:attribute-value node name))

(defun node-text (node)
  "Get text content of NODE."
  (stp:string-value node))
