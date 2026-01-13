;;;; src/types.lisp

(in-package #:cl-excel)

;;; Missing values

(defstruct (missing-value (:constructor %make-missing-value)
                          (:print-function (lambda (s stream d)
                                             (declare (ignore s d))
                                             (format stream "#<MISSING>"))))
  ;; internal struct for type identity
  )

(defparameter +missing+ (%make-missing-value)
  "Sentinel value representing a missing or empty cell in Excel.")

(defun missing-p (x)
  "True if X is the +MISSING+ sentinel."
  (eq x +missing+))

(deftype xlsx-index () '(integer 1 *))

;;; Reference Types

(defstruct (cell-ref (:constructor make-cell-ref (row col &key (abs-row nil) (abs-col nil) sheet)))
  "Represents a single cell location (e.g. A1)."
  (row 1 :type xlsx-index)
  (col 1 :type xlsx-index)
  (abs-row nil :type boolean)
  (abs-col nil :type boolean)
  (sheet nil :type (or null string)))

(defstruct (range-ref (:constructor make-range-ref (from to &key sheet)))
  "Represents a rectangular range (e.g. A1:B2)."
  (from nil :type cell-ref)
  (to nil :type cell-ref)
  (sheet nil :type (or null string)))

;;; Condition hierarchy (M1)

(define-condition xlsx-error (error)
  ()
  (:documentation "Base class for all CL-EXCEL errors."))

(define-condition xlsx-parse-error (xlsx-error)
  ((message :initarg :message :reader xlsx-error-message))
  (:report (lambda (condition stream)
             (format stream "XLSX Parse Error: ~A" (xlsx-error-message condition)))))

(define-condition sheet-missing-error (xlsx-error)
  ((name :initarg :name :reader sheet-missing-name))
  (:report (lambda (condition stream)
             (format stream "Sheet not found: ~A" (sheet-missing-name condition)))))

(define-condition invalid-range-error (xlsx-error)
  ((ref :initarg :ref :reader invalid-range-ref))
  (:report (lambda (condition stream)
             (format stream "Invalid cell/range reference: ~A" (invalid-range-ref condition)))))

(define-condition read-only-error (xlsx-error)
  ()
  (:report "Attempted to write to a read-only workbook."))

;;; Data Types (M3)

(defstruct (cell (:constructor make-cell (value &key type style-id)))
  "Represents a single cell's data and metadata."
  (value nil)      ; The actual value (string, number, date, boolean)
  (type nil)       ; :s (shared), :n (number), :b (bool), :e (error), :str (formula)
  (style-id nil))  ; Integer index into styles

(defclass sheet ()
  ((name :initarg :name :accessor sheet-name)
   (id :initarg :id :accessor sheet-id)
   (rel-id :initarg :rel-id :accessor sheet-rel-id)
   (cells :initarg :cells :accessor sheet-cells 
          :initform (make-hash-table :test 'equal)) ; Key: (cons row col)
   (dimension :initarg :dimension :accessor sheet-dimension :initform nil)
   (tables :initarg :tables :accessor sheet-tables :initform nil)) ; List of table objects
  (:documentation "In-memory representation of a worksheet."))

(defstruct table-column
  id
  name)

(defclass table ()
  ((id :initarg :id :accessor table-id)
   (name :initarg :name :accessor table-name)         ; Excel internal name (e.g. "Table1")
   (display-name :initarg :display-name :accessor table-display-name)
   (ref :initarg :ref :accessor table-ref)            ; Range string "A1:C5"
   (columns :initarg :columns :accessor table-columns); List of table-column
   (sheet :initarg :sheet :accessor table-sheet :initform nil)) ; Back-reference
  (:documentation "Represents an Excel Table (ListObject)."))

