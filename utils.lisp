;;;; utils.lisp

(in-package #:plump-utils)

;;; plump shortcuts

(defun find-child/tag (parent child-tag-name)
  (find child-tag-name (plump:children parent)
	:key #'(lambda (c) (and (plump:element-p c) (plump:tag-name c)))
	:test #'equal))

(defun find-children/tag (parent child-tag-name)
  (remove-if-not (alexandria:curry #'equal child-tag-name)
		 (coerce (plump:children parent) 'list)
		 :key #'plump:tag-name))

(defun find-child/tag/val (root tag)
  (alexandria:when-let ((element (find-child/tag root tag)))
    (plump:attribute element "w:val")))

(defun ensure-child/tag (parent child-tag-name &optional first)
  (or (find-child/tag parent child-tag-name)
      (if (not first)
	  (plump:make-element parent child-tag-name)
	  (let* ((root (plump:make-root))
		 (new-child (plump:make-element root child-tag-name)))
	    (plump:prepend-child parent new-child)))))

(defun remove-child/tag (parent child-tag-name)
  (alexandria:when-let (child (find-child/tag parent child-tag-name))
    (plump:remove-child child)))

(defun remove-if/tag (tag sequence)
  (plump:ensure-child-array
   (remove-if (alexandria:curry #'string= tag)
	      sequence
	      :key #'plump:tag-name)))

(defun remove-if-not/tag (tag sequence)
  (plump:ensure-child-array
   (remove-if-not (alexandria:curry #'string= tag)
		  sequence
		  :key #'plump:tag-name)))

(defun make-element/attrs (root tag-name &rest attributes)
  (plump:make-element root tag-name :attributes (alexandria:plist-hash-table attributes :test 'equalp)))

(defun get-first-element-by-tag-name (node tag)
  (labels ((scan-children (node)
	     (loop for child across (plump:children node)
		   when (plump:element-p child)
			do (when (tagp child tag)
			     (return-from get-first-element-by-tag-name child))
			   (scan-children child))))
    (scan-children node)
    nil))

(defun tagp (node tag)
  (string= tag (plump:tag-name node)))

(defun find-ancestor-element (node ancestor-tag)
  (loop for next = (plump:parent node) then (plump:parent next)
	while (plump:child-node-p next)
	when (tagp next ancestor-tag)
	  do (return next)
	finally (return nil)))	

(defun preservep (string)
  (or (alexandria:starts-with #\Space string :test #'char=)
      (alexandria:ends-with #\Space string :test #'char=)))

(defun make-text-element (parent text node-name)
  (let ((wt (plump:make-element parent node-name)))
    (when (preservep text)
      (setf (plump:attribute wt "xml:space") "preserve"))
    (plump:make-text-node wt text)
    wt))

(defmacro increment-attribute (element attribute)
  (alexandria:once-only (element attribute)
    `(setf (plump:attribute ,element ,attribute)
	   (princ-to-string (1+ (parse-integer (plump:attribute ,element ,attribute)))))))
