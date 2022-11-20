;;;; workbook-writer.lisp

(in-package #:sml)

(defclass workbook-writer ()
  ((%document
    :accessor workbook-writer-document)
   (%sst
    :accessor workbook-writer-sst)
   (%sheets
    :initform nil
    :accessor workbook-writer-sheets)
   (%dynamic-array-formulas
    :initform nil
    :accessor workbook-has-dynamic-array-formulas)))

(defun make-workbook-writer ()
  (let ((ww (make-instance 'workbook-writer))
	(doc (make-document)))
    (add-workbook doc)
    (add-styles doc) 
    (add-shared-string-table doc)
    (setf (workbook-writer-document ww) doc
	  (workbook-writer-sst ww) (make-string-table))
    ww))

(defun write-workbook (ww filename)
  (prepare-to-write ww)
  (docxplora:save-document (workbook-writer-document ww) filename))

(defun prepare-to-write (ww)
  (prepare-string-table ww)
  (prepare-worksheets ww)
  (when (workbook-has-dynamic-array-formulas ww)
    (ensure-dynamic-metadata ww)))

(defun ensure-dynamic-metadata (ww)
  (let* ((md-element (get-first-element-by-tag-name
		     (opc:xml-root
		      (ensure-metadata (workbook-writer-document ww)))
		     "metadata"))
	 (md-type-ref (ensure-dynamic-metadata-type md-element))
	 (cm-ref (ensure-dynamic-cell-metadata md-element md-type-ref)))
    (ensure-dynamic-future-metadata md-element)
    (setf (workbook-has-dynamic-array-formulas ww) cm-ref)))

(defun ensure-dynamic-metadata-type (md-element)
  (alexandria:if-let ((mt (get-first-element-by-tag-name md-element "metadataTypes")))
    (alexandria:if-let ((xladpr-entry (find "XLDAPR" (plump:children mt)
					    :key (lambda (entry) (plump:attribute entry "name"))
					    :test #'string-equal)))
      (princ-to-string (plump:child-position xladpr-entry))
      (let ((next (princ-to-string (1+ (parse-integer (plump:attribute mt "count"))))))
	(make-xladpr-entry mt)
	(setf (plump:attribute mt "count") next)
	next))
    (prog1
	"1"
      (make-xladpr-entry
       (make-element/attrs md-element "metadataTypes" "count" "1")))))

(defun ensure-dynamic-future-metadata (md-element)
  (let ((xladpr-fm (find "XLADPR" (plump:get-elements-by-tag-name md-element "futureMetadata")
			 :key (lambda (entry) (plump:attribute entry "name"))
			 :test #'string-equal)))
    (unless xladpr-fm
      (let ((previous-element (or (car (plump:get-elements-by-tag-name md-element "futureMetadata"))
				  (get-first-element-by-tag-name md-element "metadataTypes")))
	    (new-fm
	      (make-element/attrs
	       (make-element/attrs
		(make-element/attrs
		 (make-element/attrs
		  (make-element/attrs (plump:make-root) "futureMetadata" "name" "XLDAPR" "count" "1")
		  "bk")
		 "extLst")
		"ext" "uri" "{bdbb8cdc-fa1e-496e-a857-3c3f30c029c3}")
	       "xda:dynamicArrayProperties" "fDynamic" "1" "fCollapsed" "0")))
	(plump:insert-after previous-element new-fm)))))

(defun ensure-dynamic-cell-metadata (md-element md-type-ref)
  (let ((cm (get-first-element-by-tag-name md-element "cellMetadata")))
    (if cm
	(let ((new-count (princ-to-string (1+ (parse-integer (plump:attribute cm "count"))))))
	  (make-dynamic-cell-metadata-entry cm md-type-ref)
	  (setf (plump:attribute cm "count") new-count))
	(make-dynamic-cell-metadata-entry
	 (make-element/attrs md-element "cellMetadata" "count" "1")
	 md-type-ref))))

(defun make-dynamic-cell-metadata-entry (parent metadata-type-ref)
  (make-element/attrs
   (make-element/attrs parent "bk")
   "rc"
   "t" metadata-type-ref
   "v" "0"))
		      

(defun make-xladpr-entry (parent)
  (make-element/attrs
   parent
   "metadataType"
   "name" "XLDAPR"
   "minSupportedVersion" "120000"
   "copy" "1"
   "pasteAll" "1"
   "pasteValues" "1"
   "merge" "1"
   "splitFirst" "1"
   "rowColShift" "1"
   "clearFormats" "1"
   "clearComments" "1"
   "assign" "1"
   "coerce" "1"
   "cellMeta" "1"))
  
(defun fill-sst (string-table sst)
  (let ((words (mapcar #'car (sort (alexandria:hash-table-alist (sst-dict string-table)) #'< :key #'cdr))))
    (setf (plump:attribute sst "count") (princ-to-string (sst-count string-table))
	  (plump:attribute sst "uniqueCount") (princ-to-string (sst-unique-count string-table)))
    (loop for word in words
	  for si = (plump:make-element sst "si")
	  if (char= #\uFFFF (char word 0))
	    do (opc:parse (subseq word 1) :root si)
	  else
	    do (make-text-element si word "t"))
    sst))

(defgeneric prepare-string-table (ww)
  (:method ((ww workbook-writer))
    (let ((string-table (workbook-writer-sst ww))
	  (sst-element (get-first-element-by-tag-name (opc:xml-root (shared-string-table (workbook-writer-document ww))) "sst")))
      (fill-sst string-table sst-element))))

(defclass ww-sheet ()
  ((%ww
    :initarg :workbook-writer
    :accessor workbook-writer)
   (%name
    :initarg :name
    :accessor ww-sheet-name)
   (%table
    :initarg :table
    :accessor ww-sheet-table)))

(defmethod add-worksheet ((ww workbook-writer) &optional name)
  (ww-add-sheet ww name))

(defun ww-add-sheet (ww &optional name)
  (let ((sheet
	  (make-instance 'ww-sheet :workbook-writer ww
				   :name name
				   :table (make-cell-table))))
    (push sheet (workbook-writer-sheets ww))
    sheet))

(defun add-dimension (root sheet)
  (multiple-value-bind (row-min col-min row-max col-max)
      (table-dimensions (ww-sheet-table sheet))
    (make-element/attrs root "dimension" "ref" (encode-a1-range row-min col-min row-max col-max))))

(defun add-sheet-data (root sheet)
  (make-sheet-data (ww-sheet-table sheet) root))

(defgeneric prepare-worksheets (ww)
  (:method ((ww workbook-writer))
    (let ((doc (workbook-writer-document ww)))
      (dolist (ws (workbook-writer-sheets ww))
	(let* ((ws-part (add-worksheet doc (ww-sheet-name ws)))
	       (ws-element (get-first-element-by-tag-name (opc:xml-root ws-part) "worksheet")))
	  (prepare-worksheet ws-element ws))))))

(defun prepare-worksheet (root sheet)
  (add-dimension root sheet)
  (add-sheet-data root sheet))

;;; Writing cells

(defun serialize-rich-text (node)
  (concatenate 'string (string #\uFFFF) (opc:serialize node nil)))

(defun write-cell* (sheet ref object)
  (multiple-value-bind (row col)
      (decode-a1 ref)
    (write-cell sheet row col object)))

(defgeneric write-cell (sheet row col object)
  (:method ((sheet ww-sheet) (row integer) (col integer) (string string))
    (let ((index (string-index (workbook-writer-sst (workbook-writer sheet)) (string-xstring string))))
      (insert-cell (ww-sheet-table sheet) (list :string index) row col)))
  (:method ((sheet ww-sheet) (row integer) (col integer) (node plump:node))
    (let ((index (string-index (workbook-writer-sst (workbook-writer sheet))
			       (serialize-rich-text node))))
      (insert-cell (ww-sheet-table sheet) (list :string index) row col)))
  (:method ((sheet ww-sheet) (row integer) (col integer) (boolean (eql t)))
    (insert-cell (ww-sheet-table sheet) (list :boolean boolean) row col))
  (:method ((sheet ww-sheet) (row integer) (col integer) (boolean (eql nil)))
    (insert-cell (ww-sheet-table sheet) (list :boolean boolean) row col))
  (:method ((sheet ww-sheet) (row integer) (col integer) (number real))
    (insert-cell (ww-sheet-table sheet) (list :number number) row col))
  (:method ((sheet ww-sheet) (row integer) (col integer) (complex complex))
    (insert-cell (ww-sheet-table sheet) (list :complex complex) row col))
  (:method ((sheet ww-sheet) (row integer) (col integer) (timestamp local-time:timestamp))
    (insert-cell (ww-sheet-table sheet) (list :timestamp timestamp) row col)))

(defun write-rich-text-inline* (sheet ref text)
  (multiple-value-bind (row col)
      (decode-a1 ref)
    (write-rich-text-inline sheet row col text)))

(defun write-rich-text-inline (sheet row col text)
  (insert-cell (ww-sheet-table sheet) (list :rich-text text) row col))

(defun write-formula* (sheet ref formula)
  (multiple-value-bind (row col)
      (decode-a1 ref)
    (write-formula sheet row col formula)))

(defun write-formula (sheet row col formula)
  (insert-cell (ww-sheet-table sheet) (list :formula (fix-formula formula)) row col))

(defparameter *dynamic-array-formula-scanner*
  (cl-ppcre:create-scanner
   "\\b(?:SORT|LET|LAMBDA|SINGLE|SORTBY|UNIQUE|XMATCH|FILTER|XLOOKUP|SEQUENCE|RANDARRAY|ANCHORARRAY)\\("
   :case-insensitive-mode t))

(defun rename-dynamic-array-formula (formula)
  (dolist (entry *dynamic-renamings* formula)
    (setf formula (cl-ppcre:regex-replace-all (first entry) formula (second entry)))))

(defun rename-future-formula (formula)
  (dolist (entry *future-renamings* formula)
    (setf formula (cl-ppcre:regex-replace-all (first entry) formula (second entry)))))

(defun fix-formula (formula)
  (rename-dynamic-array-formula (rename-future-formula (string-left-trim "=" formula))))

(defun dynamic-array-formula-p (formula)
  (cl-ppcre:scan *dynamic-array-formula-scanner* formula))

(defun write-array-formula (sheet start-row start-col end-row end-col formula &aux (type :array-formula))
  (when (dynamic-array-formula-p formula)
    (setf type :dynamic-array-formula
	  (workbook-has-dynamic-array-formulas (workbook-writer sheet)) t))
  (insert-cell (ww-sheet-table sheet) (list type
					    (fix-formula formula)
					    end-row end-col)
	       start-row start-col)
  (loop for row from start-row to end-row
	do (loop for col from start-col to end-col
		 do (unless (and (= col start-col)
				 (= row start-row))
		      (insert-cell (ww-sheet-table sheet) (list :number 0) row col)))))

(defun write-array-formula* (sheet ref formula)
  (destructuring-bind (start end)
      (serapeum:split-sequence #\: ref)
    (multiple-value-bind (start-row start-col)
	(decode-a1 start)
      (multiple-value-bind (end-row end-col)
	  (decode-a1 end)
	(write-array-formula sheet start-row start-col end-row end-col formula)))))

(defun write-row (sheet row col sequence)
  (serapeum:do-each (item sequence)
    (write-cell sheet row col item)
    (incf col)))

(defun write-row* (sheet ref sequence)
  (multiple-value-bind (row col)
      (decode-a1 ref)
    (write-row sheet row col sequence)))

(defun write-column (sheet row col sequence)
  (serapeum:do-each (item sequence)
    (write-cell sheet row col item)
    (incf row)))

(defun write-column* (sheet ref sequence)
  (multiple-value-bind (row col)
      (decode-a1 ref)
    (write-column sheet row col sequence)))

(defun write-array (sheet row col array)
  (loop for ar below (array-dimension array 0)
	for tr from row
	do (loop for ac below (array-dimension array 1)
		 for tc from col
		 do (write-cell sheet tr tc (aref array ar ac)))))

(defun write-array* (sheet ref array)
  (multiple-value-bind (row col)
      (decode-a1 ref)
    (write-array sheet row col array)))
