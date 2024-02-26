;;;; tables.lisp

(cl:in-package #:docxplora)

(defun tables-in-document-order (node)
  (let ((result '()))
    (plump:traverse
     node
     #'(lambda (node) (when (tagp node "w:tbl") (push node result)))
     :test #'plump:element-p)
    (nreverse result)))

(defun table-rows (table)
  (find-children/tag table "w:tr"))

(defun table-row-cells (table-row)
  (find-children/tag table-row "w:tc"))

(defun table-all-cells (table)
  (loop for row across (table-rows table)
	appending (table-row-cells row)))

(serapeum:define-do-macro do-rows ((row table &optional return) &body body)
  `(map nil (lambda (,row)
	      ,@body)
	(table-rows ,table)))
