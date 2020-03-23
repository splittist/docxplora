;;;; recolor.lisp

(cl:defpackage #:recolor
  (:use #:cl)
  (:export
   #:recolor-string))

(cl:in-package #:recolor)

(defgeneric process-node (node))

(defmethod process-node (node)
  (let ((tag (plump:tag-name node))
	(class (plump:attribute node "class")))
    (if class
	(cond
	  ((string= "default" class)
	   nil)
	 ((serapeum:string^= "paren" class)
	  nil)
	 (t (format t "~A: " class)))
	(format t "<~A>:" tag))))

(defmethod process-node ((node plump-dom:text-node))
  (format t "\"~A\"" (plump:text node)))

(defparameter *style* '())

(defgeneric handle-node (node))

(defmethod handle-node ((node plump-dom:root))
  (serapeum:do-each (child (plump:children node))
    (handle-node child)))

(defmethod handle-node (node)
  (let ((tag (plump:tag-name node))
	(class (plump:attribute node "class"))
	(children (plump:children node)))
    (if class
	(cond
	  ((string= "" class)
	   (setf class nil))
	  ((string= "default" class)
	   (setf class nil))
	  ((serapeum:string^= "paren" class)
	   (setf class nil))
	  (t
	   nil))
	(setf class tag))
    (when class (push class *style*))
    (serapeum:do-each (child children)
      (handle-node child))
    (when class (pop *style*))))

(defparameter *escape-table*
  (serapeum:dictq
   #\< "&lt;"
   #\> "&gt;"
   #\& "&amp;"
   #\" "&quot;"
   #\' "&apos;"))

(defun preservep (string)
  (or (alexandria:starts-with #\Space string :test #'char=)
      (alexandria:ends-with #\Space string :test #'char=)))

(defun write-text (line)
  (let ((parts (split-sequence:split-sequence #\Tab line)))
    (flet ((write-part (part)
	     (write-string "<w:t")
	     (when (preservep part) (write-string " xml:space=\"preserve\""))
	     (write-string ">")
	     (write-string part)
	     (write-string "</w:t>")))
      (let ((first-part (pop parts)))
	(unless (alexandria:emptyp first-part)
	  (write-part first-part)))
      (dolist (part parts)
	(write-string "<w:tab/>")
	(unless (alexandria:emptyp part)
	  (write-part part))))))

(defmethod handle-node ((node plump-dom:text-node))
  (let* ((text (serapeum:escape (plump:text node) *escape-table*))
	 (lines (split-sequence:split-sequence #\Newline text)))
    (write-string "<w:r>")
    (when *style*
      (format t "<w:rPr><w:rStyle w:val=\"md~A\"/>~@[<w:i/>~]</w:rPr>" (first *style*)(second *style*)))
    (let ((first-line (pop lines)))
      (write-text first-line))
    (dolist (line lines)
      (write-string "<w:br/>")
      (unless (alexandria:emptyp line)
	(write-text line)))
    (write-string "</w:r>")))

(defun recolor-string (string &optional (coloring-type :lisp))
  (let ((root (plump:parse (colorize:html-colorization coloring-type string))))
    (with-output-to-string (s)
      (let ((*standard-output* s))
	(handle-node root)))))

(defun test (string &optional (coloring-type :lisp))
  (let* ((a (recolor-string string coloring-type))
	 (b (with-output-to-string (s)
	      (plump:serialize (plump:parse a) s))))
    (let ((result (mismatch a b)))
      (if result
	  (values result (subseq a result))
	  t))))
