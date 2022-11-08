;;;; table.lisp

(in-package #:sml)

(defclass table (sml-xml-part)
  ())

(defgeneric tables (target)
  (:method ((worksheet worksheet))
    (let* ((table-refs (plump:get-elements-by-tag-name (opc:xml-root worksheet) "tablePart"))
	   (ids (mapcar #'(lambda (trf) (plump:attribute trf "r:id")) table-refs)))
      (mapcar #'(lambda (id)
		  (let* ((rel (opc:get-relationship worksheet id))
			 (target-uri (opc:target-uri rel))
			 (source-uri (opc:source-uri rel))
			 (abs-uri (opc:uri-merge source-uri target-uri)))
		    (docxplora:get-part-by-name worksheet abs-uri :xml t :class 'table)))
	      ids)))
  (:method ((document sml-document))
    (alexandria:mappend #'tables (worksheets document))))

(defun next-table-uri (document)
  (let* ((sheets (worksheets document))
	 (rels (alexandria:mappend (lambda (sheet) (opc:get-relationships-by-type-code sheet "TABLE"))
				   sheets))
	 (targets (mapcar #'opc:target-uri rels)))
    (loop for i from 1
	  for candidate = (format nil "../tables/table~D.xml" i)
	  while (find candidate targets :test 'string-equal)
	  finally (return candidate))))

(defgeneric add-table (document worksheet)
  (:method ((document sml-document) (worksheet worksheet))
    (let* ((package (docxplora:opc-package document))
	   (wb (workbook document))
	   (ws-tables (get-first-element-by-tag-name (opc:xml-root worksheet) "tableParts"))
	   (uri (next-table-uri document))
	   (part (create-sml-xml-part
		  'table
		  package
		  uri
		  (opc:ct "SML_TABLE")))
	   (xml-root (opc:xml-root part))
	   (root (plump:make-element xml-root "table")))
      (dolist (nss (list nil "mc" "xr" "xr3"))
	(let ((ns (assoc nss *workbook-namespaces* :test #'equal)))
	  (setf (plump:attribute root (format nil "xmlns~@[:~A~]" (car ns))) (cdr ns))))
      (setf (plump:attribute root "mc:Ignorable") "xr xr3")
      (let ((reln (opc:create-relationship wb
					   (opc:uri-relative (opc:part-name worksheet) uri)
					   (opc:rt "TABLE"))))
	(make-element/attrs ws-tables "tablePart"
			    "r:id" (opc:relationship-id reln)))
      part)))
