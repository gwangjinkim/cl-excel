;;;; src/styles.lisp

(in-package #:cl-excel)

(defstruct styles
  (num-fmts (make-hash-table :test 'eql)) ;; Map numFmtId (int) -> formatCode (string)
  (cell-xfs #()))                         ;; Vector of numFmtIds (ints), indexed by xfId

(defun read-styles (zip)
  "Read xl/styles.xml from ZIP. Returns a STYLES struct."
  (let ((stream (get-entry-stream zip "xl/styles.xml"))
        (styles (make-styles :num-fmts (make-hash-table :test 'eql)
                             :cell-xfs #())))
    (if stream
        (let* ((dom (parse-xml stream))
               (root (stp:document-element dom))
               ;; 1. Parse <numFmts>
               (num-fmts-node (find-child root "numFmts")))
          
          (when num-fmts-node
            (stp:do-children (nf num-fmts-node)
              (when (and (typep nf 'stp:element) (string= (stp:local-name nf) "numFmt"))
                (let ((id (parse-integer (get-attribute nf "numFmtId")))
                      (code (get-attribute nf "formatCode")))
                  (setf (gethash id (styles-num-fmts styles)) code)))))
          
          ;; 2. Parse <cellXfs> (ordered list)
          (let ((xfs-node (find-child root "cellXfs"))
                (xf-list '()))
            (when xfs-node
              (stp:do-children (xf xfs-node)
                (when (and (typep xf 'stp:element) (string= (stp:local-name xf) "xf"))
                  (let ((num-fmt-id (parse-integer (get-attribute xf "numFmtId"))))
                    (push num-fmt-id xf-list)))))
            (setf (styles-cell-xfs styles) (coerce (nreverse xf-list) 'vector)))
          
          styles)
        ;; If no styles, return empty struct
        styles)))

(defun is-date-format (fmt-code)
  "Heuristic check if format code represents a date/time."
  (when fmt-code
    ;; Common date tokens: y, m, d, h, s, :, /, -
    ;; Avoid false positives like "General", "0.00", "Red"
    (let ((lower (string-downcase fmt-code)))
      (or (search "yy" lower)
          (search "mm" lower)
          (search "dd" lower)
          (search "hh" lower)
          (search "ss" lower)
          (search "m/d" lower)
          (search "d-m" lower)))))

(defun get-style-num-fmt (styles xf-id)
  "Get the numFmtId for a given xf index."
  (when (and styles (< xf-id (length (styles-cell-xfs styles))))
    (aref (styles-cell-xfs styles) xf-id)))
