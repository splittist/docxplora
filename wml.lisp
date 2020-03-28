;;;; wml.lisp

(cl:in-package #:docxplora)

(defun add-main-document (document)
  (let* ((package (opc-package document))
	 (part (opc:create-xml-part package "/word/document.xml" (opc:ct "WML_DOCUMENT_MAIN")))
	 (root (opc:xml-root part))
	 (doc (plump:make-element root "w:document"))
	 (body (plump:make-element doc "w:body")))
    (plump:make-element body "w:p")
    (dolist (ns *md-namespaces*)
      (setf (plump:attribute doc (format nil "xmlns:~A" (car ns))) (cdr ns)))
    (dolist (mc *md-ignorable*)
      (setf (plump:attribute doc (car mc)) (cdr mc)))
    (opc:create-relationship package "/word/document.xml" (opc:rt "OFFICE_DOCUMENT"))
    part))

(defun ensure-main-document (document)
  (alexandria:if-let ((existing (main-document document)))
    existing
    (add-main-document document)))

;;; FIXME can also be target of glossary-document

(defun md-target (document rt)
  (let* ((main-document (main-document document))
	 (rel (first (opc:get-relationships-by-type main-document rt))))
    (when rel
      (let ((target (opc:uri-merge (opc:part-name main-document) (opc:target-uri rel))))
	(get-part-by-name document target t)))))

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

(defun add-numbering-definitions (document)
  (let* ((package (opc-package document))
	 (md (main-document document))
	 (part (opc:create-xml-part package "/word/numbering.xml" (opc:ct "WML_NUMBERING")))
	 (root (opc:xml-root part))
	 (numbering (plump:make-element root "w:numbering")))
    (dolist (ns *md-namespaces*)
      (setf (plump:attribute numbering (format nil "xmlns:~A" (car ns))) (cdr ns)))
    (dolist (mc *md-ignorable*)
      (setf (plump:attribute numbering (car mc)) (cdr mc)))
    (opc:create-relationship md (opc:uri-relative "/word/document.xml" "/word/numbering.xml") (opc:rt "NUMBERING"))
    part))

(defun style-definitions (document)
  (md-target document (opc:rt "STYLES")))

(defun add-style-definitions (document)
  (let* ((package (opc-package document))
	 (md (main-document document))
	 (part (opc:create-xml-part package "/word/styles.xml" (opc:ct "WML_STYLES")))
	 (root (opc:xml-root part))
	 (styles (plump:make-element root "w:styles")))
    (dolist (nss (list "mc" "r" "w" "w14" "w15" "w16cid" "w16se"))
      (let ((ns (assoc nss *md-namespaces* :test #'string=)))
	(setf (plump:attribute styles (format nil "xmlns:~A" (car ns))) (cdr ns))))
    (setf (plump:attribute styles "mc:Ignorable") "w14 w15 w16cid w16se") ; FIXME - better way to set markup compatibilty
    (opc:create-relationship md (opc:uri-relative "/word/document.xml" "/word/styles.xml") (opc:rt "STYLES"))
    part))

(defun ensure-style-definitions (document)
  (alexandria:if-let ((sd (style-definitions document)))
    sd
    (add-style-definitions document)))

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
		       (abs-uri (opc:uri-merge source-uri target-uri)))
		  (get-part-by-name document abs-uri t)))
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

