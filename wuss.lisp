;;;; wuss.lisp

(cl:in-package #:wuss)

(defun read-forms (stream)
  (uiop:with-safe-io-syntax ()
    (uiop:slurp-stream-forms stream)))

(defclass source ()
  ((%items :initarg :items :initform () :accessor source-items)))

(defmethod print-object ((object source) stream)
  (print-unreadable-object (object stream :type t :identity t)
    (write-string (serapeum:ellipsize (princ-to-string (source-items object)) 12) stream)))

(defun make-source (&optional items)
  (make-instance 'source :items items))

(defgeneric peek-source (source)
  (:method ((source source))
    (if (null (source-items source))
	(values nil nil)
	(values (car (source-items source)) t))))

(defgeneric pop-source (source)
  (:method ((source source))
    (if (null (source-items source))
	(values nil nil)
	(values (pop (source-items source)) t))))

(defgeneric push-source (source item)
  (:method ((source source) item)
    (push item (source-items source))))

(defun to-camelcase (symbol &key upper-first special-words)
  (let ((parts (split-sequence:split-sequence #\- (string symbol))))
    (apply
     #'serapeum:string+
     (loop for part in parts
	for firstp = t then nil
	for special = (member part special-words :test #'string-equal)
	collect (cond ((and firstp (not upper-first))
		       (string-downcase part))
		      (special (first special))
		      (t (string-capitalize part)))))))

(defun frob-form (form)
  (typecase form
    (null nil)
    (symbol (to-camelcase form))
    (string form)
    (t (princ-to-string form))))

(defun slurp-items (source)
  (let ((attrs '())
	(children '())
	(tag (pop-source source)))
    (loop for (next left) = (multiple-value-list (peek-source source))
       until (or (null left)
		 (keywordp next))
       do (if (consp next)
	      (push (pre-process (make-source (pop-source source))) children)
	      (push (pop-source source) attrs))
	 finally (return `(,tag ,(nreverse attrs) ,(nreverse children))))))

(defun pre-process (source)
  (loop for (next left) = (multiple-value-list (peek-source source))
     while left
     collecting
       (etypecase next
	 (keyword (slurp-items source))
	 (cons (pre-process (pop-source source))))))

(defun process (forms)
  (dolist (form forms)
    (render-form (first form) (second form) (third form))))

(defgeneric render-form (tag attrs children)
  (:method (tag attrs children)
    (when (oddp (length attrs))
      (setf attrs (cons "val" attrs)))
    (format t "<w:~A ~{w:~A=\"~A\" ~}>" (frob-form tag) (mapcar #'frob-form attrs))
    (dolist (child children) (process child))
    (format t "</w:~A>" (frob-form tag))))
      
