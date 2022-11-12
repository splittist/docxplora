;;;; endnotes.lisp

(in-package #:docxplora)

(defun get-endnote-by-id (root id)
  (when (numberp id) (setf id (princ-to-string id)))
  (let ((endnotes (plump:get-elements-by-tag-name root "w:endnote")))
    (find-if (lambda (e) (equal id (plump:attribute e "w:id"))) endnotes)))

(defun find-endnote-property-value (root property)
  (alexandria:when-let ((endnote-properties (find-child/tag root "w:endnotePr")))
    (find-child/tag/val endnote-properties property)))

(defun applicable-endnote-reference-numbering-value (property default endnote-reference sections document)
  (or (find-endnote-property-value (endnote-reference-section endnote-reference sections) property)
      (find-child/tag/val (find-document-setting document "w:endnotePr") property)
      default))

(defun applicable-endnote-reference-numbering-start (endnote-reference sections document)
  (applicable-endnote-reference-numbering-value "w:numStart" "1" endnote-reference sections document))

(defun applicable-endnote-reference-numbering-restart (endnote-reference sections document)
  (applicable-endnote-reference-numbering-value "w:numRestart" "continuous" endnote-reference sections document))

(defun applicable-endnote-reference-numbering-format (endnote-reference sections document)
  ;;; Note 17.11.17 says "decimal", which seems wrong
  (applicable-endnote-reference-numbering-value "w:numFmt" "lowerRoman" endnote-reference sections document))

(defun endnote-references-in-document-order (node)
  (let ((result '())
	(paragraphs (paragraphs-in-document-order node)))
    (dolist (paragraph paragraphs)
      (plump:traverse
       paragraph
       (lambda (node) (when (tagp node "w:endnoteReference") (push node result)))
       :test #'plump:element-p))
    (nreverse result)))

(defgeneric endnotes-in-document-order (target)
  (:method ((document wml-document))
    (let* ((md-root (opc:xml-root (main-document document)))
	   (fn-root (opc:xml-root (endnotes document)))
	   (endnote-references (endnote-references-in-document-order md-root))
	   (endnote-ids (mapcar #'(lambda (ref) (plump:attribute ref "w:id")) endnote-references))
	   (endnotes (plump:get-elements-by-tag-name fn-root "w:endnote")))
      (mapcar #'(lambda (id) (find-if #'(lambda (f) (equal id (plump:attribute f "w:id"))) endnotes))
	      endnote-ids))))

(defun endnote-reference-paragraph (endnote-reference)
  (loop for node = endnote-reference then (plump:parent node)
	while node
	when (equal "w:p" (plump:tag-name node))
	  do (return-from endnote-reference-paragraph node)
	finally (return nil)))

(defun endnote-reference-section (endnote-reference sections)
  (paragraph-section (endnote-reference-paragraph endnote-reference) sections))

(defun endnote-references-numbering (document)
  (let* ((md-root (opc:xml-root (main-document document)))
	 (refs (endnote-references-in-document-order md-root))
	 (sections (sections document)))
    (loop with previous-section = nil
	  with fmt = nil
	  with restart = nil
	  with count = 1
	  for ref in refs
	  for custom = (on-off-attribute ref "w:customMarkFollows")
	  for section = (endnote-reference-section ref sections)
	  unless (eql section previous-section)
	    do (setf previous-section section
		     fmt (applicable-endnote-reference-numbering-format ref sections document)
		     restart (applicable-endnote-reference-numbering-restart ref sections document))
	       ;;; The MS Word GUI constrains this to "eachSect" and "continuous", which makes sense
	       (when (equal "eachSect" restart)
		 (setf count (parse-integer (applicable-endnote-reference-numbering-start ref sections document))))
	  collecting (list (if (not custom) (format-number fmt count) "")
			   ref)
	  unless custom do (incf count))))
