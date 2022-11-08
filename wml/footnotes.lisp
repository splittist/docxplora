;;;; footnotes.lisp

(in-package #:docxplora)

(defun get-footnote-by-id (root id)
  (when (numberp id) (setf i (princ-to-string id)))
  (let ((footnotes (plump:get-elements-by-tag-name root "w:footnote")))
    (find-if (lambda (f) (equal id (plump:attribute f "w:id"))) footnotes)))

(defun find-footnote-property-value (root property)
  (alexandria:when-let ((footnote-properties (find-child/tag root "w:footnotePr")))
    (find-child/tag/val footnote-properties property)))

(defun applicable-footnote-reference-numbering-value (property default footnote-reference sections document)
  (or (find-footnote-property-value (footnote-reference-section footnote-reference sections) property)
      (find-child/tag/val (find-document-setting document "w:footnotePr") property)
      default))

(defun applicable-footnote-reference-numbering-start (footnote-reference sections document)
  (applicable-footnote-reference-numbering-value "w:numStart" "1" footnote-reference sections document))

(defun applicable-footnote-reference-numbering-restart (footnote-reference sections document)
  (applicable-footnote-reference-numbering-value "w:numRestart" "continuous" footnote-reference sections document))

(defun applicable-footnote-reference-numbering-format (footnote-reference sections document)
  (applicable-footnote-reference-numbering-value "w:numFmt" "decimal" footnote-reference sections document))

(defun footnote-references-in-document-order (node)
  (let ((result '()))
    (plump:traverse
     node
     #'(lambda (node) (when (tagp node "w:footnoteReference") (push node result)))
     :test #'plump:element-p)
    (nreverse result)))

(defgeneric footnotes-in-document-order (target)
  (:method ((document wml-document))
    (let* ((md-root (opc:xml-root (main-document document)))
	   (fn-root (opc:xml-root (footnotes document)))
	   (footnote-references (footnote-references-in-document-order md-root))
	   (footnote-ids (mapcar #'(lambda (ref) (plump:attribute ref "w:id")) footnote-references))
	   (footnotes (plump:get-elements-by-tag-name fn-root "w:footnote")))
      (mapcar #'(lambda (id) (find-if #'(lambda (f) (equal id (plump:attribute f "w:id"))) footnotes))
	      footnote-ids))))

(defun footnote-reference-paragraph (footnote-reference)
  (loop for node = footnote-reference then (plump:parent node)
	while node
	when (equal "w:p" (plump:tag-name node))
	  do (return-from footnote-reference-paragraph node)
	finally (return nil)))

(defun footnote-reference-section (footnote-reference sections)
  (paragraph-section (footnote-reference-paragraph footnote-reference) sections))

(defun footnote-references-numbering (document)
  (let* ((md-root (opc:xml-root (main-document document)))
	 (refs (footnote-references-in-document-order md-root))
	 (sections (sections document)))
    (loop with previous-section = nil
	  with fmt = nil
	  with restart = nil
	  with count = 1
	  for ref in refs
	  for custom = (on-off-attribute ref "w:customMarkFollows")
	  for section = (footnote-reference-section ref sections)
	  unless (eql section previous-section)
	    do (setf previous-section section
		     fmt (applicable-footnote-reference-numbering-format ref sections document)
		     restart (applicable-footnote-reference-numbering-restart ref sections document))
	       (when (equal "eachSect" restart) ;; FIXME - eachPage
		 ;;; MS Word does something differently that effectively adds the numStart to the running count
		 ;;; when a new section with a different START is encountered, even if the number is CONTINUOUS 
		 (setf count (parse-integer (applicable-footnote-reference-numbering-start ref sections document))))
	  collecting (list (if (not custom) (format-number fmt count) "")
			   ref)
	  unless custom do (incf count))))
