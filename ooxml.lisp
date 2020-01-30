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
	  (plump:parse (flexi-streams:octets-to-string (opc:content part) :external-format :utf8))))
  part)

(defgeneric get-part-by-name (document name &optional ensure-xml)
  (:method ((document document) (name string) &optional ensure-xml)
    (let* ((package (opc-package document))
	   (part (opc:get-part package name)))
      (if ensure-xml
	  (ensure-xml part)
	  part))))
    
(defun main-document (document)
  (let* ((package (opc-package document))
	 (rel (first (opc:get-relationships-by-type package (opc:rt "OFFICE_DOCUMENT"))))
	 (target (opc:uri-merge "/" (opc:target-uri rel))))
    (get-part-by-name document target t)))

;;; FIXME can also be target of glossary-document

(defun md-target (document rt)
  (let* ((main-document (main-document document))
	 (rel (first (opc:get-relationships-by-type main-document rt))))
    (when rel
      (let((target (opc:uri-merge (opc:part-name main-document) (opc:target-uri rel))))
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

;;; styles

(defun make-style-definitions (document) ;; FIXME Not Yet Implemented
  ;; make-instance; add xml; get right name; add to package; add to relationships, creating if necessary; return part
  nil)
  
(defgeneric add-style (target style)
  (:method ((document document) (style plump-dom:element))
    (unless (equal "w:style" (plump:tag-name style))
      (error "Need a \"w:style\" element to add to document; got ~A" style))
    (let* ((style-definitions
	    (or (style-definitions document)
		(make-style-definitions document)))
	   (styles (first (plump:get-elements-by-tag-name
			   (opc:xml-root style-definitions)
			   "w:styles"))))
      (plump:append-child styles style))))

(defgeneric remove-style (target style)
  (:method ((document document) (style-id string))
    (alexandria:when-let (style-element (find-style-by-id document style-id))
      (plump:remove-child style-element)))
  (:method ((document document) (style plump:element))
    (plump:remove-child style)))

(defgeneric find-style-by-id (target style-id &optional include-latent) ;; FIXME mixes ids and names
  (:method ((document document) (style-id string) &optional include-latent)
    (alexandria:when-let ((style-definitions (style-definitions document)))
      (let ((styles (first (plump:get-elements-by-tag-name
			    (opc:xml-root style-definitions)
			    "w:styles"))))
	(or (find-if (alexandria:curry #'equal style-id)
		     (plump:get-elements-by-tag-name styles "w:style")
		     :key (alexandria:rcurry #'plump:attribute "w:styleId"))
	    (and include-latent
		 (find-if (alexandria:curry #'equal style-id)
			  (plump:get-elements-by-tag-name styles "w:lsdException")
			  :key (alexandria:rcurry #'plump:attribute "w:name"))))))))

;;; utils

(defun find-child/tag (parent child-tag-name)
  (find child-tag-name (plump:children parent) :key #'plump:tag-name :test #'equal))

(defun find-children/tag (parent child-tag-name)
  (remove-if-not (alexandria:curry #'equal child-tag-name)
		 (coerce (plump:children parent) 'list)
		 :key #'plump:tag-name))

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

(defun make-element/attrs (root tag-name &rest attributes)
  (plump:make-element root tag-name :attributes (alexandria:plist-hash-table attributes :test 'equalp)))
