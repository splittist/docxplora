;;;; numbering.lisp

(cl:in-package #:docxplora)

(defun get-next-num-id (numbering-definitions)
  (let* ((root (opc:xml-root numbering-definitions))
	 (nums (plump:get-elements-by-tag-name root "w:num"))
	 (digits (mapcar (alexandria:rcurry #'plump:attribute "w:numId") nums)))
    (loop for i from 1
       while (find (princ-to-string i) digits :test #'string=)
	 finally (return i))))

(defun get-next-abstract-num-id (numbering-definitions)
  (let* ((root (opc:xml-root numbering-definitions))
	 (nums (plump:get-elements-by-tag-name root "w:abstractNum"))
	 (digits (mapcar (alexandria:rcurry #'plump:attribute "w:abstractNumId") nums)))
    (loop for i from 1
       while (find (princ-to-string i) digits :test #'string=)
	 finally (return i))))

(defun make-numbering-start-override (numbering-definitions abstract-num-id &key (ilvl 0) (start 1))
  (let ((num-id (get-next-num-id numbering-definitions)))
    (plump:first-child
     (plump:parse
      (wuss:compile-style
       `(:num num-id ,num-id
	 (:abstract-num-id ,abstract-num-id
	  (:lvl-override ilvl ,ilvl     
	   (:start-override ,start)))))))))

(defun add-numbering-definition (numbering-definitions numbering-definition)
  (let ((numbering (get-first-element-by-tag-name (opc:xml-root numbering-definitions)
						  "w:numbering")))
    (plump:append-child numbering numbering-definition)))

