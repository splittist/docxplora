(cl:in-package #:docxplora)

(defclass document ()
  ((%package :initarg :package :accessor opc-package)))

(defun open-document (pathname)
  (let ((opc-package (opc:open-package pathname)))
    (make-instance 'document :package opc-package)))

(defun ensure-xml (part)
  (when (eql 'opc:opc-part (type-of part))
    (change-class part 'opc:opc-xml-part)
    (setf (opc:xml-root part)
	  (plump:parse (flexi-streams:octets-to-string (opc:content part))))))

(defun main-document (document)
  (let* ((package (opc-package document))
	 (rel (first (opc:get-relationships-by-type package (opc:rt "OFFICE_DOCUMENT"))))
	 (target (opc:uri-merge "/" (opc:target-uri rel)))
	 (mdp (opc:get-part package target)))
    (ensure-xml mdp)
    mdp))

;;; FIXME can also be target of glossary-document

(defun md-target (document rt)
  (let* ((package (opc-package document))
	 (main-document (main-document document))
	 (rel (first (opc:get-relationships-by-type main-document rt))))
    (when rel
      (let* ((target (opc:uri-merge (opc:part-name main-document) (opc:target-uri rel)))
	     (part (opc:get-part package target)))
	(ensure-xml part)
	part))))

(defun comments (document)
  (md-target document (opc:rt "COMMENTS")))

(defun document-settings (document)
  (md-target document (opc:rt "SETTINGS")))

(defun endnotes (document)
  (md-target document (opc:rt "ENDNOTES")))

(defun footnotes (document)
  (md-target document (opc:rt "FOOTNOTES")))

(defun font-table (document)
  (md-target document (opc:rt "FONT_TABLE")))

(defun glossary-document (document)
  (md-target document (opc:rt "GLOSSARY_DOCUMENT")))

(defun numbering-definitions (document)
  (md-target document (opc:rt "NUMBERING")))

(defun style-definitions (document)
  (md-target document (opc:rt "STYLES")))

(defun web-settings (document)
  (md-target document (opc:rt "WEB_SETTINGS")))


(defun md-internal-targets (document ref)
  (let* ((mdp (main-document document))
	 (header-refs (plump:get-elements-by-tag-name (opc:xml-root mdp) ref))
	 (ids (mapcar #'(lambda (hrf) (plump:attribute hrf "r:id")) header-refs)))
    (mapcar #'(lambda (id)
		(let* ((rel (opc:get-relationship mdp id))
		       (target-uri (opc:target-uri rel))
		       (source-uri (opc:source-uri rel))
		       (abs-uri (opc:uri-merge source-uri target-uri))
		       (part (opc:get-part (opc-package document) abs-uri)))
		  (ensure-xml part)
		  part))
	    ids)))

(defun headers (document)
  (md-internal-targets document "w:headerReference"))

(defun footers (document)
  (md-internal-targets document "w:footerReference"))

;; Document template - explicit External relationship of Document Settings part @ w:attachedTemplate via r:id with
;;  .../attachedTemplate relationship type

;; Framesets External docx files tageted from WebSettings part with .../frame rel type

;; Master Document main-document targets External docx @ w:subDoc via r:id with .../SubDocument rel type

;; Mail Merge Data Source explicit External relationship of document-settings @ w:dataSource via r:id
;;  with ../mailMergeSource rel type

;; XSL Transformation ditto @ w:saveThroughXslt via r:id with .../transform

