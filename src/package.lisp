;;;; src/package.lisp — packages and exports (M0)

(defpackage #:cl-excel
  (:use #:cl)
  (:nicknames #:xlsx)
  (:export
   ;; Open/close
   #:readxlsx #:openxlsx #:closexlsx #:with-xlsx #:writexlsx
   ;; Workbook/sheets
   #:sheetnames #:sheetcount #:hassheet #:sheet #:addsheet! #:rename!
   ;; Cells/ranges
   #:getdata #:readdata #:getcell #:getcellrange
   ;; Indexing sugar (optional API surface)
   #:cell #:range #:sheetref #:named
   ;; Iterators
   #:eachrow #:row-number #:eachtablerow
   ;; Tables
   #:readtable #:gettable #:datatable-data #:datatable-column-labels
   #:datatable-column-index #:writetable #:writetable!))

