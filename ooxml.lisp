;;;; ooxml.lisp

(cl:in-package #:docxplora)

(defclass document ()
  ((%package :initarg :package :accessor opc-package)))

(defun open-document (pathname)
  (let ((opc-package (opc:open-package pathname)))
    (make-instance 'document :package opc-package)))

(defun make-document ()
  (let ((opc-package (make-instance 'opc:opc-package)))
    (make-instance 'document :package opc-package)))

(defun save-document (document &optional pathname)
  (opc:save-package (opc-package document) pathname))

(defgeneric get-part-by-name (document name &optional ensure-xml)
  (:method ((document document) (name string) &optional ensure-xml)
    (let* ((package (opc-package document))
	   (part (opc:get-part package name)))
      (if ensure-xml
	  (opc:ensure-xml part)
	  part))))
    
(defun main-document (document)
  (let* ((package (opc-package document))
	 (rel (first (opc:get-relationships-by-type package (opc:rt "OFFICE_DOCUMENT"))))
	 (target (opc:uri-merge "/" (opc:target-uri rel))))
    (get-part-by-name document target t)))

(defun document-type (document) ; FIXME - return actual types
  (let ((mdp-ct (opc:content-type (main-document document))))
    (cond ((member mdp-ct *wml-content-types* :test #'string-equal)
	   :wordprocessing-document)
	  ((member mdp-ct *sml-content-types* :test #'string-equal)
	   :spreadsheet-document)
	  ((member mdp-ct *pml-content-types* :test #'string-equal)
	   :presentation-document)
	  (t :opc-package))))

;;; WML

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

;;; Hyperlinks (external)

(defun find-hyperlink (source uri)
  (find uri (opc:get-relationships-by-type source (opc:rt "HYPERLINK"))
	:key #'opc:target-uri
	:test #'string-equal)) ; FIXME Genrally check string= / string-equal

(defun ensure-hyperlink (source uri)
  (alexandria:if-let ((existing (find-hyperlink source uri)))
    existing
    (opc:create-relationship source uri (opc:rt "HYPERLINK") :target-mode "External")))

;;; styles

(defun make-style-definitions (document) ; FIXME Not Yet Implemented
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

(defgeneric find-style-by-id (target style-id &optional include-latent) ; FIXME mixes ids and names
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

(defun coalesce-adjacent-text (run)
  (let ((groups (serapeum:runs (plump:children run) :key #'plump:tag-name :test #'string=)))
    (dolist (group groups)
      (when (and (< 1 (length group))
		 (string= "w:t" (plump:tag-name (alexandria:first-elt group))))
	(let ((first-node (alexandria:first-elt group))
	      (text (serapeum:string-join (map 'list #'plump:text group))))
	  (unless (plump:first-child first-node)
	    (plump:make-text-node first-node)) ; FIXME test if this can create nonsense
	  (setf (plump:text (plump:first-child first-node)) text) ; textual node
	  (when (char= #\Space (alexandria:last-elt text))
	    (setf (plump:attribute first-node "xml:space") "preserve"))
	  (serapeum:do-each (node (subseq group 1))
	    (plump:remove-child node)))))))

#||

Style elements:

__General__

basedOn - val= styleID of parent

link - val=styleID ; links character to paragraph styles and vice versa



name - val=string - primary style name

next - val=styleId ; style for next paragraph; paragraph styles only

qFormat - present or absent - treated as a primary style by application



aliases - val= comma separated list of style names (for UI)

autoRedefine - present or absent (interaction)

hidden - present or absent (UI)

locked - present or absent (UI)

personal - val=boolean (email message context); character styles only

personalCompose - val=boolean (email message context); character styles only

personalReply - ditto

rsid - val=four byte number (Long Hexidecimal Number Value) - editing session

semiHidden - present or absent - initially visible (UI)

uiPriority - val=number optional UI sorting order

unhideWhenUsed - present or absent



__Latent Styles__

latentStyles - behaviour properties (rather than formatting properties) for application (UI etc)
  attributes: count=number, defLockedState=boolean, defQFormat=boolean, defSemiHidden=boolean, defUIPriority=number, defUnhideWhenUsed=boolean, 

lsdException - exception for named latent style
  attributes: name=primary-name (as known to application), locked=boolean, qFormat=boolean, semiHidden=boolean, uiPriority=number, unhideWhenUsed=boolean



__styles.xml__

styles - container for style definitions and latent styles

w:docDefaults>w:rPrDefault>w:rPr^w:pPrDefault>w:pPr

w:latentStyles goes here

style - general style properties; style type (paragraph, character, table, numbering); style type-specific properties
  attributes: customStyle=boolean; default=boolean; styleId=string; type="paragraph" etc

__Character Styles__

type="character"

General Elements

w:rPr

__Paragraph Styles__

type="paragraph"

General Elements

w:rPr - for all runs

w:pPr - paragraph properties; numbering reference is ref to numbering definition only, which in turn has a reference to the paragraph styleon the level associated with the style

__Table Styles__

type="table"

General Elements

w:tblPr

w:tblStylePr w:type="firstRow"|... ; inlcudes w:tblPr {heaps of stuff}

__Numbering Styles__

type="numbering"

w:pPr>w:numPr>w:numId[w:val="~D"

only info specified is ref to numbering defintion

__Run Properties__

b - toggle (bold)

bCs - toggle (complex script bold)

bdo - val="rtl"|"ltr"

bdr - [17.3.2.5]


||#
