;;;; workbook-editor.lisp

(in-package #:sml)

(defclass workbook-editor (workbook-writer)
  ((%document
    :accessor workbook-editor-document)
   (%sst
    :accessor workbook-editor-sst
    :initform nil)
   (%edited-sheets
    :initform nil
    :accessor workbook-editor-edited-sheets)
   (%sheets
    :accessor workbook-editor-added-sheets)))

(defun make-workbook-editor (path)
  (let ((we (make-instance 'workbook-editor))
	(document (open-document path)))
    (setf (workbook-editor-document we) document
	  (workbook-editor-sst we) (read-sst (shared-string-table document)))
    (opc:delete-part (docxplora:opc-package document) "/xl/calcChain.xml")
    we))

(defgeneric edit-worksheet (workbook-editor name)
  (:method ((we workbook-editor) (name string))
    (let* ((worksheet (get-worksheet-by-name (workbook-editor-document we) name))
	   (sheet (make-instance 'ww-sheet :workbook-writer we
					   :name name
					   :table (read-sheet-data worksheet))))
      (push sheet (workbook-editor-edited-sheets we))
      sheet)))

(defun replace-string-table (we)
  (let ((string-table (workbook-editor-sst we))
	(sst-element (get-first-element-by-tag-name
		      (opc:xml-root (shared-string-table (workbook-editor-document we)))
		      "sst")))
    (loop for child across (plump:children sst-element)
	  do (plump:remove-child child))
    (fill-sst string-table sst-element)))

(defmethod prepare-worksheets :after ((we workbook-editor))
  (prepare-edited-worksheets we))

(defun prepare-edited-worksheets (we)
  (dolist (ws (workbook-editor-edited-sheets we))
    (let* ((ws-part (get-worksheet-by-name (workbook-editor-document we) (ww-sheet-name ws)))
	   (ws-element (get-first-element-by-tag-name (opc:xml-root ws-part) "worksheet")))
      (prepare-edited-worksheet ws-element ws))))

(defun prepare-edited-worksheet (ws-element ws)
  (update-dimension ws-element ws)
  (replace-sheet-data ws-element ws))

(defun update-dimension (ws-element ws)
  (multiple-value-bind (row-min col-min row-max col-max)
      (table-dimensions (ww-sheet-table ws))
    (let ((dimension (get-first-element-by-tag-name ws-element "dimension")))
      (setf (plump:attribute dimension "ref")
	    (encode-a1-range row-min col-min row-max col-max)))))

(defun replace-sheet-data (ws-element ws)
  (let ((sd (get-first-element-by-tag-name ws-element "sheetData")))
    (loop for child across (plump:children sd) do (plump:remove-child child))
    (let ((new-sd (make-sheet-data (ww-sheet-table ws) sd)))
      (plump:splice new-sd))))

(defmethod get-worksheet-by-name ((we workbook-editor) (name string))
  (get-worksheet-by-name (workbook-editor-document we) name))

(defun read-sheet-data (worksheet)
  (let* ((cell-table (make-cell-table))
	 (root (opc:xml-root worksheet))
	 (sd-element (get-first-element-by-tag-name root "sheetData")))
    (loop for row across (plump:children sd-element)
	  for row-num = (parse-integer (plump:attribute row "r"))
	  for cell-table-row = (ensure-row cell-table row-num)
	  do (setf (cell-table-row-attributes cell-table-row)
		   (alexandria:copy-hash-table (plump:attributes row)))
	     (loop for cell across (plump:children row)
		   for cell-ref = (plump:attribute cell "r")
		   do (multiple-value-bind (cell-row-num col-num)
			  (decode-a1 cell-ref)
			(assert (= row-num cell-row-num))
			(trees:insert (cons col-num cell) (cell-table-row-tree cell-table-row)))))
    cell-table))

(defun read-sst (shared-string-table)
  (let* ((sst (make-string-table))
	 (root (opc:xml-root shared-string-table))
	 (sst-element (get-first-element-by-tag-name root "sst"))
	 (unique-count (parse-integer (plump:attribute sst-element "uniqueCount")))
	 (sis '()))
    (plump:traverse (opc:xml-root shared-string-table)
		    #'(lambda (node) (when (tagp node "si") (push node sis)))
		    :test #'plump:element-p)
    (dolist (si (nreverse sis))
      (let ((child (plump:first-child si)))
	(if (tagp child "t")
	    (string-index sst (plump:text child))
	    (string-index
	     sst
	     (with-output-to-string (s)
	       (princ #\uFFFF s)
	       (loop for child across (plump:children si)
		     do (opc:serialize child s)))))))
    (unless (= unique-count (sst-unique-count sst))
      (warn "Unique counts do not match: read ~S, wanted ~S"
	    unique-count (sst-unique-count sst)))
    sst))
