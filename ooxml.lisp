;;;; ooxml.lisp

(cl:in-package #:docxplora)

(defclass document ()
  ((%package :initarg :package :accessor opc-package)))

;; moved to wml to create wml-documents
#+(or)(defun open-document (pathname)
  (let ((opc-package (opc:open-package pathname)))
    (make-instance 'document :package opc-package)))

#+(or)(defun make-document ()
  (let ((opc-package (make-instance 'opc:opc-package)))
    (make-instance 'document :package opc-package)))

(defun save-document (document &optional pathname)
  (opc:save-package (opc-package document) pathname))

(defgeneric get-part-by-name (document name &key xml class)
  (:method ((document document) (name string) &key xml class)
    (let* ((package (opc-package document))
	   (part (opc:get-part package name)))
      (when xml (opc:ensure-xml part))
      (if class (change-class part class) part))))
    
(defgeneric main-document (document)
  (:method (document)
    (let* ((package (opc-package document))
	   (rel (first (opc:get-relationships-by-type package (opc:rt "OFFICE_DOCUMENT"))))
	   (target (opc:uri-merge "/" (opc:target-uri rel))))
      (get-part-by-name document target :xml t))))

(defun document-type (document) ; FIXME - return actual types
  (let ((mdp-ct (opc:content-type (main-document document))))
    (cond ((member mdp-ct *wml-content-types* :test #'string-equal)
	   :wordprocessing-document)
	  ((member mdp-ct *sml-content-types* :test #'string-equal)
	   :spreadsheet-document)
	  ((member mdp-ct *pml-content-types* :test #'string-equal)
	   :presentation-document)
	  (t :opc-package))))

