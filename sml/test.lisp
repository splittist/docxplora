;;;; test.lisp

(in-package #:sml)

(defparameter *test-heading* '("Region" "Sales Rep" "Product" "Units"))

(defparameter *test-data*
  #2A(("East"  "Tom"    "Apple"   6380)
      ("West"  "Fred"   "Grape"   5619)
      ("North" "Amy"    "Pear"    4565)
      ("South" "Sal"    "Banana"  5323)
      ("East"  "Fritz"  "Apple"   4394)
      ("West"  "Sravan" "Grape"   7195)
      ("North" "Xi"     "Pear"    5231)
      ("South" "Hector" "Banana"  2427)
      ("East"  "Tom"    "Banana"  4213)
      ("West"  "Fred"   "Pear"    3239)
      ("North" "Amy"    "Grape"   6420)
      ("South" "Sal"    "Apple"   1310)
      ("East"  "Fritz"  "Banana"  6274)
      ("West"  "Sravan" "Pear"    4894)
      ("North" "Xi"     "Grape"   7580)
      ("South" "Hector" "Apple"   9814)))

(defun filter-test (outpath)
  (let* ((ww (make-workbook-writer))
	 (ws (add-worksheet ww)))
    ;; Sales transacton table
    (write-cell* ws :a2 "Sales Transactions")
    (write-row* ws :a4 *test-heading*)
    (write-array* ws :a5 *test-data*)
    ;; Filter and results
    (write-cell* ws :g2 "Product:")
    (write-row* ws :f4 *test-heading*)
    (write-cell* ws :h2 "Apple")
    (write-array-formula* ws "F5:I8" "FILTER(A5:D20,C5:C20=H2)")
    ;; Write workbook
    (write-workbook ww outpath)))

(defparameter *test-rich-text*
  (opc:parse "<r><rPr><b/><color rgb=\"FFBC8F8F\"/><sz val=\"24\"/><u val=\"double\"/></rPr><t>foo</t></r><r><t>bar</t></r>"))

(defun rich-text-test (outpath)
  (let* ((ww (make-workbook-writer))
	 (ws (add-worksheet ww)))
    (write-rich-text-inline ws 2 2 *test-rich-text*)
    (write-cell ws 2 3 " <= Inline")
    (write-cell ws 3 3 *test-rich-text*)
    (write-cell ws 3 2 "SST => ")
    (write-workbook ww outpath)))

(defun edit-test (outpath)
  (let* ((we (make-workbook-editor "c:/Users/David/Downloads/styles.xlsx"))
	 (ws (edit-worksheet we "Sheet1")))
    (write-cell* ws :c3 "Jan")
    (write-workbook we outpath)))
 
