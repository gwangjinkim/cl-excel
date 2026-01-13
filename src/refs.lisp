;;;; src/refs.lisp

(in-package #:cl-excel)

;;; Column conversions

(defun col-name (index)
  "Convert 1-based column index to Excel column string (e.g. 1 -> A, 28 -> AB)."
  (unless (and (integerp index) (plusp index))
    (error 'invalid-range-error :ref index))
  (let ((n index)
        (chars '()))
    (loop while (> n 0)
          do (let ((rem (mod (1- n) 26)))
               (push (code-char (+ (char-code #\A) rem)) chars)
               (setf n (floor (1- n) 26))))
    (coerce chars 'string)))

(defun col-index (name)
  "Convert Excel column string to 1-based index (e.g. A -> 1, AB -> 28)."
  (unless (and (stringp name) (plusp (length name)))
    (error 'invalid-range-error :ref name))
  (let ((sum 0))
    (loop for char across (string-upcase name)
          do (let ((val (1+ (- (char-code char) (char-code #\A)))))
               (unless (and (>= val 1) (<= val 26))
                 (error 'invalid-range-error :ref name))
               (setf sum (+ (* sum 26) val))))
    sum))

;;; Parsing



(defun parse-range (ref-string)
  "Parse a range string (e.g. 'A1:B2' or 'Sheet1!A1:B2') into a RANGE-REF."
  (let ((colon-pos (position #\: ref-string)))
    (unless colon-pos
      (error 'invalid-range-error :ref ref-string))
    (let* ((part1 (subseq ref-string 0 colon-pos))
           (part2 (subseq ref-string (1+ colon-pos)))
           (from (parse-cell-ref part1))
           (to (parse-cell-ref part2)))
      (make-range-ref from to :sheet (cell-ref-sheet from)))))

(defun parse-cell-ref (ref-string &key range)
  "Parse a cell reference or range string."
  (if (and range (find #\: ref-string))
      (parse-range ref-string)
      (parse-cell-ref-impl ref-string))) ;; Call original implementation (renamed)

(defun parse-cell-ref-impl (ref-string)
 "Internal implementation of single cell parsing."
   (let ((sheet nil)
         (rest ref-string))
    ;; Handle Sheet! prefix
    (let ((bang-pos (position #\! ref-string :from-end t)))
      (when bang-pos
        (setf sheet (subseq ref-string 0 bang-pos))
        (setf rest (subseq ref-string (1+ bang-pos)))))
    
    ;; Simple state machine to separate letters and digits
    (let ((col-str (make-array 0 :element-type 'character :fill-pointer 0 :adjustable t))
          (row-str (make-array 0 :element-type 'character :fill-pointer 0 :adjustable t))
          (abs-col nil)
          (abs-row nil)
          (phase :start)) ; :start -> :col -> :row
      
      (loop for i from 0 below (length rest)
            for char = (char rest i)
            do (cond
                 ((char= char #\$)
                  (if (eq phase :start)
                      (setf abs-col t)
                      (setf abs-row t)))
                 ((alpha-char-p char)
                  (if (eq phase :row)
                      (error 'xlsx-parse-error :message (format nil "Invalid char ~C in row part" char))
                      (progn
                        (setf phase :col)
                        (vector-push-extend char col-str))))
                 ((digit-char-p char)
                  (setf phase :row)
                  (vector-push-extend char row-str))
                 (t (error 'xlsx-parse-error :message (format nil "Unexpected char ~C in ref" char)))))
      
      (when (zerop (length col-str))
        (error 'xlsx-parse-error :message "Missing column part"))
      (when (zerop (length row-str))
        (error 'xlsx-parse-error :message "Missing row part"))
      
      (make-cell-ref (parse-integer row-str)
                     (col-index col-str)
                     :abs-row abs-row
                     :abs-col abs-col
                     :sheet sheet))))

;;; Range Accessors Helpers

(defun range-start-row (r) (cell-ref-row (range-ref-from r)))
(defun range-end-row (r) (cell-ref-row (range-ref-to r)))
(defun range-start-col (r) (cell-ref-col (range-ref-from r)))
(defun range-end-col (r) (cell-ref-col (range-ref-to r)))

;;; Formatting

(defun cell-ref-to-string (ref)
  "Convert a CELL-REF struct back to a string (e.g. A1)."
  (with-output-to-string (s)
    (when (cell-ref-sheet ref)
      (format s "~A!" (cell-ref-sheet ref)))
    (when (cell-ref-abs-col ref) (write-char #\$ s))
    (write-string (col-name (cell-ref-col ref)) s)
    (when (cell-ref-abs-row ref) (write-char #\$ s))
    (format s "~D" (cell-ref-row ref))))
