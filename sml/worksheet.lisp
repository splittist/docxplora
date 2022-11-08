;;;; worksheet.lisp

(in-package #:sml)

(defclass worksheet (sml-xml-part)
  ())

(defgeneric worksheets (document)
  (:method ((document sml-document))
    (let* ((wb (workbook document))
	   (sheet-refs (plump:get-elements-by-tag-name (opc:xml-root wb) "sheet"))
	   (ids (mapcar #'(lambda (shrf) (plump:attribute shrf "r:id")) sheet-refs)))
      (mapcar #'(lambda (id)
		  (let* ((rel (opc:get-relationship wb id))
			 (target-uri (opc:target-uri rel))
			 (source-uri (opc:source-uri rel))
			 (abs-uri (opc:uri-merge source-uri target-uri)))
		    (docxplora:get-part-by-name document abs-uri :xml t :class 'worksheet)))
	      ids))))

(defgeneric get-worksheet-by-name (target name)
  (:method ((document sml-document) (name string))
    (let* ((wb (workbook document))
	   (sheet-refs (plump:get-elements-by-tag-name (opc:xml-root wb) "sheet"))
	   (sheet-ref (find name sheet-refs :key (lambda (shrf) (plump:attribute shrf "name"))
					    :test #'string-equal)))
      (when sheet-ref
	(let* ((rel (opc:get-relationship wb (plump:attribute sheet-ref "r:id")))
	       (target-uri (opc:target-uri rel))
	       (source-uri (opc:source-uri rel))
	       (abs-uri (opc:uri-merge source-uri target-uri)))
	  (docxplora:get-part-by-name document abs-uri :xml t :class 'worksheet))))))

(defun check-name (wb name) ;; FIXME check for empty name etc.
  (let* ((sheet-refs (plump:get-elements-by-tag-name (opc:xml-root wb) "sheet"))
	 (sheet-ids (mapcar #'(lambda (shrf) (plump:attribute shrf "sheetId")) sheet-refs))
	 (names (mapcar #'(lambda (shrf) (plump:attribute shrf "name")) sheet-refs))
	 (id-num (1+ (loop for sheet-id in sheet-ids
			   maximizing (parse-integer sheet-id))))
	 (uri (format nil "/xl/worksheets/sheet~D.xml" id-num)))
    (if name
	(if (find name names :test #'string=) ;; FIXME case sensitive?
	    (error "Duplicate sheet name: ~A" name) ;; FIXME conditions
	    (values name uri id-num))
	(if (null sheet-refs)
	    (values "Sheet1" uri id-num)
	    (values (format nil "Sheet~D" id-num) uri id-num)))))

(defgeneric add-worksheet (document &optional name)
  (:method ((document sml-document) &optional name)
    (let* ((package (docxplora:opc-package document))
	   (wb (workbook document))
	   (wb-sheets (get-first-element-by-tag-name (opc:xml-root wb) "sheets")))
      (multiple-value-bind (name uri sheet-id)
	  (check-name wb name)
	(let* ((part (create-sml-xml-part
		      'worksheet
		      package
		      uri
		      (opc:ct "SML_WORKSHEET")))
	       (xml-root (opc:xml-root part))
	       (root (plump:make-element xml-root "worksheet")))
	  (dolist (nss (list nil "r" "mc" "x14ac" "xr" "xr2" "xr3"))
	    (let ((ns (assoc nss *workbook-namespaces* :test #'equal)))
	      (setf (plump:attribute root (format nil "xmlns~@[:~A~]" (car ns))) (cdr ns))))
	  (setf (plump:attribute root "mc:Ignorable") "x14ac xr xr2 xr3")
	  (let ((reln (opc:create-relationship wb (opc:uri-relative "/xl/workbook.xml" uri)
					       (opc:rt "WORKSHEET"))))
	    (make-element/attrs wb-sheets "sheet"
				"name" name
				"sheetId" (princ-to-string sheet-id)
				"r:id" (opc:relationship-id reln)))
	  part)))))


(defun make-sheet-data (cell-table &optional parent)
  (let ((spans (table-spans cell-table))
	(sd (plump:make-element (or parent (plump:make-root)) "sheetData")))
    (trees:dotree (row-entry (cell-table-rows cell-table))
      (let* ((row-num (car row-entry))
	     (row-obj (cdr row-entry))
	     (span-index (floor row-num 16))
	     (span (nth span-index spans))
	     (row-element (plump:make-element sd "row"
					      :attributes (cell-table-row-attributes row-obj))))
	(setf (plump:attribute row-element "r") (princ-to-string row-num)
	      (plump:attribute row-element "spans") (format nil "~D:~D" (car span) (cdr span)))
	(trees:dotree (cell-entry (cell-table-row-tree row-obj))
	  (handle-cell-element row-element row-num (car cell-entry) (cdr cell-entry)))))
    sd))

(defgeneric handle-cell-element (row-element row-num col-num cell-data)
  (:method (row-element row-num col-num (cell plump:element))
    (plump:append-child row-element (plump:clone-node cell)))
  (:method (row-element row-num col-num (cell-data cons))
    (make-cell-element row-num col-num row-element (first cell-data) (rest cell-data))))

(defgeneric make-cell-element (row col parent type rest)
  (:method (row col parent (type (eql :string)) rest)
    (let* ((cell-element (make-element/attrs parent "c" "r" (encode-a1 row col) "t" "s"))
	   (value-element (plump:make-element cell-element "v")))
      (plump:make-text-node value-element (princ-to-string (first rest)))))
  (:method (row col parent (type (eql :boolean)) rest)
    (let* ((cell-element (make-element/attrs parent "c" "r" (encode-a1 row col) "t" "b"))
	   (value-element (plump:make-element cell-element "v")))
      (plump:make-text-node value-element (if (first rest) "1" "0"))))
  (:method (row col parent (type (eql :number)) rest)
    (let* ((cell-element (make-element/attrs parent "c" "r" (encode-a1 row col)))
	   (value-element (plump:make-element cell-element "v")))
      (plump:make-text-node value-element (string-right-trim " ." (format nil "~,,,,,,'EG" (first rest))))))
  (:method (row col parent (type (eql :complex)) rest)
    (let* ((cell-element (make-element/attrs parent "c" "r" (encode-a1 row col) "t" "f"))
	   (formula-element (plump:make-element cell-element "f"))
	   (value-element (plump:make-element cell-element "v")))
      (plump:make-text-node formula-element
			    (format nil "COMPLEX(~A, ~A)" (realpart (first rest)) (imagpart (first rest))))
      (plump:make-text-node value-element "0")))
  (:method (row col parent (type (eql :timestamp)) rest)
    (let* ((cell-element (make-element/attrs parent "c" "r" (encode-a1 row col)))
	   (value-element (plump:make-element cell-element "v")))
      (plump:make-text-node value-element (prin1-to-string (timestamp-excel (first rest))))))
  (:method (row col parent (type (eql :formula)) rest)
    (let* ((cell-element (make-element/attrs parent "c" "r" (encode-a1 row col) "t" "f"))
	   (formula-element (plump:make-element cell-element "f"))
	   (value-element (plump:make-element cell-element "v")))
      (plump:make-text-node formula-element (first rest))
      (plump:make-text-node value-element "0")))
  (:method (row col parent (type (eql :array-formula)) rest)
    (destructuring-bind (formula end-row end-col)
	rest
      (let* ((ref (format nil "~A:~A" (encode-a1 row col) (encode-a1 end-row end-col)))
	     (cell-element (make-element/attrs parent "c" "r" (encode-a1 row col) "t" "f"))
	     (formula-element (make-element/attrs cell-element "f" "t" "array" "ref" ref))
	     (value-element (plump:make-element cell-element "v")))
	(plump:make-text-node formula-element formula)
	(plump:make-text-node value-element "0"))))
  (:method (row col parent (type (eql :dynamic-array-formula)) rest)
    (destructuring-bind (formula end-row end-col)
	rest
      (let* ((ref (encode-a1-range row col end-row end-col))
	     (cell-element (make-element/attrs parent "c" "r" (encode-a1 row col) "t" "f" "cm" "1"))
	     (formula-element (make-element/attrs cell-element "f" "t" "array" "ref" ref))
	     (value-element (plump:make-element cell-element "v")))
	(plump:make-text-node formula-element formula)
	(plump:make-text-node value-element "0"))))
  (:method (row col parent (type (eql :rich-text)) rest)
    (let* ((cell-element (make-element/attrs parent "c" "r" (encode-a1 row col) "t" "inlineStr"))
	   (is-element (plump:make-element cell-element "is")))
      (loop for child across (plump:children (first rest))
	    do (plump:append-child is-element child)))))

;; write rather than insert; read rather than find???

