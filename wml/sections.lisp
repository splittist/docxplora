;;;; sections.lisp

(in-package #:docxplora)

(defun last-section-properties (root)
  (let ((body (get-first-element-by-tag-name root "w:body")))
    (find-child/tag body "w:sectPr")))

(defgeneric sections (document)
  (:method ((document wml-document))
    (sections (main-document document)))
  (:method ((main-document main-document))
    (sections (opc:xml-root main-document)))
  (:method ((root plump-dom:root))
    (loop with acc = '()
	  with result = '()
	  for p in (paragraphs-in-document-order root)
	  for s = (get-first-element-by-tag-name p "w:sectPr")
	  do (push p acc)
	  when s
	    do (push (list s (nreverse acc)) result)
	       (setf acc '())
	  finally
	     (push (list (last-section-properties root) (nreverse acc)) result)
	     (return (nreverse result)))))

(defun paragraph-section (paragraph sections)
  (loop for (section paragraphs) in sections
	when (find paragraph paragraphs)
	  do (return-from paragraph-section section)
	finally (return nil)))

