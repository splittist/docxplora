;;;; cell-table.lisp

(in-package #:sml)

(defconstant +row-max+ #x100000)

(defconstant +col-max+ #x4000)

(defclass cell-table ()
  ((%cols
    :initform '()
    :accessor cell-table-cols)
   (%rows
    :initform (trees:make-binary-tree :avl #'< :key #'car)
    :accessor cell-table-rows)))

(defclass cell-table-row ()
  ((%row-attributes
    :initform (plump:make-attribute-map)
    :accessor cell-table-row-attributes)
   (%row-tree
    :initform (trees:make-binary-tree :avl #'< :key #'car)
    :accessor cell-table-row-tree)))

(defun cell-table-row-max (cell-table-row)
  (car (trees:maximum (cell-table-row-tree cell-table-row))))

(defun cell-table-row-min (cell-table-row)
  (car (trees:minimum (cell-table-row-tree cell-table-row))))

(defun make-cell-table ()
  (make-instance 'cell-table))

(defun find-row (cell-table row-num)
  (cdr (trees:find row-num (cell-table-rows cell-table))))

(defun find-cell (cell-table row-num col-num)
  (alexandria:when-let ((row (find-row cell-table row-num)))
    (cdr (trees:find col-num (cell-table-row-tree row)))))

(defun make-cell-table-row ()
  (make-instance 'cell-table-row))

(defun insert-row (cell-table row-num)
  (let ((new-row (make-cell-table-row)))
    (trees:insert (cons row-num new-row) (cell-table-rows cell-table))
    new-row))

(defun ensure-row (cell-table row-num)
  (alexandria:if-let ((result (find-row cell-table row-num)))
    result
    (insert-row cell-table row-num)))

(defun insert-cell (cell-table cell row-num col-num)
  (let ((row (ensure-row cell-table row-num)))
    (trees:insert (cons col-num cell) (cell-table-row-tree row))
    cell-table))

(defun cell-table-minimum-row (cell-table)
  (let ((rows (cell-table-rows cell-table)))
    (if (trees:emptyp rows)
	0
	(car (trees:minimum rows)))))

(defun cell-table-maximum-row (cell-table)
  (let ((rows (cell-table-rows cell-table)))
    (if (trees:emptyp rows)
	0
	(car (trees:maximum rows)))))

(defun cell-table-row-min (cell-table-row)
  (let ((row (cell-table-row-tree cell-table-row)))
    (if (trees:emptyp row)
	0
	(car (trees:minimum row)))))

(defun cell-table-row-max (cell-table-row)
  (let ((row (cell-table-row-tree cell-table-row)))
    (if (trees:emptyp row)
	0
	(car (trees:maximum row)))))

(defun cell-table-empty-p (cell-table)
  (trees:emptyp (cell-table-rows cell-table)))

(defun table-dimensions (cell-table)
  (if (cell-table-empty-p cell-table)
      (values 0 0 0 0)
      (let ((row-min (cell-table-minimum-row cell-table))
	    (row-max (cell-table-maximum-row cell-table))
	    (col-min +col-max+)
	    (col-max 0))
	(trees:dotree (row (cell-table-rows cell-table))
	  (let ((max (cell-table-row-max (cdr row)))
		(min (cell-table-row-min (cdr row))))
	    (when (> max col-max)
	      (setf col-max max))
	    (when (< min col-min)
	      (setf col-min min))))
	(values row-min col-min row-max col-max))))

(defun table-spans (cell-table)
  (let ((spans '()))
    (loop for start from 1 by 16 to (cell-table-maximum-row cell-table)
	  do (loop for row-num from start below (+ 16 start)
		   for row = (find-row cell-table row-num)
		   when row
		     maximize (cell-table-row-max row) into max
		     and
		       minimize (cell-table-row-min row) into min
		   finally (push (cons min max) spans)))
    (nreverse spans)))

(serapeum:define-do-macro do-cell-table ((cell row-num col-num cell-table &optional return) &body body)
  (alexandria:with-gensyms (row-entry cell-entry)
    `(trees:dotree (,row-entry (cell-table-rows ,cell-table))
       (let ((,row-num (car ,row-entry)))
	 (trees:dotree (,cell-entry (cell-table-row-tree (cdr ,row-entry)))
	   (let ((,col-num (car ,cell-entry))
		 (,cell (cdr ,cell-entry)))
	     ,@body))))))
