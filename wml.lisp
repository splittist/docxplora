;;;; wml.lisp

(cl:in-package #:docxplora)

(defclass wml-document (document)
  ())

(defun open-document (pathname)
  (let ((opc-package (opc:open-package pathname)))
    (make-instance 'wml-document :package opc-package)))

(defun make-document ()
  (let ((opc-package (make-instance 'opc:opc-package)))
    (make-instance 'wml-document :package opc-package)))

(defclass wml-part (opc:opc-part) ())

(defclass wml-xml-part (opc:opc-xml-part) ())

(defun create-wml-xml-part (class-name package uri content-type)
  (let ((part (opc:create-xml-part package uri content-type)))
    (change-class part class-name)))

(defclass main-document (wml-xml-part)())

(defmethod main-document ((document wml-document))
  (let* ((package (opc-package document))
	 (rel (first (opc:get-relationships-by-type package (opc:rt "OFFICE_DOCUMENT"))))
	 (target (opc:uri-merge "/" (opc:target-uri rel))))
    (get-part-by-name document target :xml t :class 'main-document)))

(defgeneric add-main-document (document)
  (:method ((document wml-document))
    (let* ((package (opc-package document))
	   (part (create-wml-xml-part
		  'main-document
		  package
		  "/word/document.xml"
		  (opc:ct "WML_DOCUMENT_MAIN")))
	   (root (opc:xml-root part))
	   (doc (plump:make-element root "w:document"))
	   (body (plump:make-element doc "w:body")))
      (plump:make-element body "w:p")
      (dolist (ns *md-namespaces*)
	(setf (plump:attribute doc (format nil "xmlns:~A" (car ns))) (cdr ns)))
      (dolist (mc *md-ignorable*)
	(setf (plump:attribute doc (car mc)) (cdr mc)))
      (opc:create-relationship package "/word/document.xml" (opc:rt "OFFICE_DOCUMENT"))
      part)))
  
(defun ensure-main-document (document)
  (alexandria:if-let ((existing (main-document document)))
    existing
    (add-main-document document)))

;;; FIXME can also be target of glossary-document

(defun md-target (document rt class)
  (let* ((main-document (main-document document))
	 (rel (first (opc:get-relationships-by-type main-document rt))))
    (when rel
      (let ((target (opc:uri-merge (opc:part-name main-document) (opc:target-uri rel))))
	(get-part-by-name document target :xml t :class class)))))

(defmacro define-xml-part (class uri content-type root-element namespaces ignorables relationship-type)
  (let ((add-name (alexandria:symbolicate 'add- class))
	(ensure-name (alexandria:symbolicate 'ensure- class))
	(pred-name (alexandria:symbolicate class (if (find #\- (symbol-name class) :test #'char=) '-p 'p))))
  `(progn
     (defclass ,class (wml-xml-part)())
     (defgeneric ,class (document)
       (:method ((document wml-document))
	 (md-target document (opc:rt ,relationship-type) ',class)))
     (defgeneric ,add-name (document)
       (:method ((document wml-document))
	 (let* ((package (opc-package document))
		(md (main-document document))
		(part (create-wml-xml-part
		       ',class
		       package
		       ,uri
		       (opc:ct ,content-type)))
		(xml-root (opc:xml-root part))
		(root (plump:make-element xml-root ,root-element)))
	   (dolist (nss ',namespaces)
	     (let ((ns (assoc nss *md-namespaces* :test #'string=)))
	       (setf (plump:attribute root (format nil "xmlns:~A" (car ns))) (cdr ns))))
	   (setf (plump:attribute root "mc:Ignorable") ,ignorables)
	   (opc:create-relationship md (opc:uri-relative "/word/document.xml" ,uri)
				    (opc:rt ,relationship-type))
	   part)))
     (defgeneric ,ensure-name (document)
       (:method ((document document))
	 (alexandria:if-let ((result (,class document)))
	   result
	   (,add-name document))))
     (defgeneric ,pred-name (thing)
       (:method (thing) nil)
       (:method ((thing ,class)) t)))))

;; FIXME w16cex

(define-xml-part comments
    "/word/comments.xml"
  "WML_COMMENTS"
  "w:comments"
  ("wpc" "cx" "cx1" "cx2" "cx3" "cx4" "cx5" "cx6" "cx7" "cx8" "mc" "aink"
	 "am3d" "o" "r" "m" "v" "wp14" "wp" "w10" "w" "w14" "w15" "w16cid" "w16"
	 "w16se" "wpg" "wpi" "wne" "wps")
  "w14 w15 w16se w16cid w16 wp14"
  "COMMENTS")

(define-xml-part endnotes
    "/word/endnotes.xml"
  "WML_ENDNOTES"
  "w:endnotes"
  ("wpc" "cx" "cx1" "cx2" "cx3" "cx4" "cx5" "cx6" "cx7" "cx8" "mc" "aink"
	 "am3d" "o" "r" "m" "v" "wp14" "wp" "w10" "w" "w14" "w15" "w16cid" "w16"
	 "w16se" "wpg" "wpi" "wne" "wps")
  "w14 w15 w16se w16cid w16 wp14"
  "ENDNOTES")

(define-xml-part footnotes
    "/word/footnotes.xml"
  "WML_FOOTNOTES"
  "w:footnotes"
  ("wpc" "cx" "cx1" "cx2" "cx3" "cx4" "cx5" "cx6" "cx7" "cx8" "mc" "aink"
	 "am3d" "o" "r" "m" "v" "wp14" "wp" "w10" "w" "w14" "w15" "w16cid" "w16"
	 "w16se" "wpg" "wpi" "wne" "wps")
  "w14 w15 w16se w16cid w16 wp14"
  "FOOTNOTES")

(define-xml-part font-table
    "/word/fontTable.xml"
  "WML_FONT_TABLE"
  "w:endnotes"
  ("mc" "r" "w" "w14" "w15" "w16cid" "w16" "w16se")
  "w14 w15 w16se w16cid w16"
  "FONT_TABLE")

(defclass glossary-document (wml-part)()) ; FIXME - what is it?

(defun glossary-document (document)
  (md-target document (opc:rt "GLOSSARY_DOCUMENT") 'glossary-document))

(define-xml-part numbering-definitions
    "/word/numbering.xml"
  "WML_NUMBERING"
  "w:numbering"
  ("wpc" "cx" "cx1" "cx2" "cx3" "cx4" "cx5" "cx6" "cx7" "cx8" "mc" "aink"
	 "am3d" "o" "r" "m" "v" "wp14" "wp" "w10" "w" "w14" "w15" "w16cid" "w16"
	 "w16se" "wpg" "wpi" "wne" "wps")
  "w14 w15 w16se w16cid w16 wp14"
  "NUMBERING")

(define-xml-part document-settings
    "/word/settings.xml"
  "WML_SETTINGS"
  "w:settings"
  ("mc" "aink" "o" "r" "m" "v" "w10" "w" "w14" "w15" "w16cid" "w16"
	 "w16se" "s1")
  "w14 w15 w16se w16cid w16"
  "SETTINGS")

(define-xml-part style-definitions
    "/word/styles.xml"
  "WML_STYLES"
  "w:styles"
  ("wpc" "cx" "cx1" "cx2" "cx3" "cx4" "cx5" "cx6" "cx7" "cx8" "mc" "aink"
	 "am3d" "o" "r" "m" "v" "wp14" "wp" "w10" "w" "w14" "w15" "w16cid" "w16"
	 "w16se" "wpg" "wpi" "wne" "wps")
  "w14 w15 w16se w16cid w16 wp14"
  "STYLES")

(define-xml-part web-settings
    "/word/webSettings.xml"
  "WML_WEB_SETTINGS"
  "w:webSettings"
  ("mc" "r" "w" "w14" "w15" "w16cid" "w16" "w16se")
  "w14 w15 w16se w16cid w16"
  "WEB_SETTINGS")

(defun md-internal-targets (document ref class)
  (let* ((mdp (main-document document))
	 (header-refs (plump:get-elements-by-tag-name (opc:xml-root mdp) ref))
	 (ids (mapcar #'(lambda (hrf) (plump:attribute hrf "r:id")) header-refs)))
    (mapcar #'(lambda (id)
		(let* ((rel (opc:get-relationship mdp id))
		       (target-uri (opc:target-uri rel))
		       (source-uri (opc:source-uri rel))
		       (abs-uri (opc:uri-merge source-uri target-uri)))
		  (get-part-by-name document abs-uri :xml t :class class)))
	    ids)))

(defclass header (wml-xml-part)())

(defun headers (document)
  (md-internal-targets document "w:headerReference" 'header))

(defclass footer (wml-xml-part)())

(defun footers (document)
  (md-internal-targets document "w:footerReference" 'footer))

;; Document template - explicit External relationship of Document Settings part @ w:attachedTemplate via r:id with
;;  .../attachedTemplate relationship type

;; Framesets External docx files tageted from WebSettings part with .../frame rel type

;; Master Document main-document targets External docx @ w:subDoc via r:id with .../SubDocument rel type

;; Mail Merge Data Source explicit External relationship of document-settings @ w:dataSource via r:id
;;  with ../mailMergeSource rel type

;; XSL Transformation ditto @ w:saveThroughXslt via r:id with .../transform

(defgeneric sections (document)
  (:method ((document wml-document))
    (let* ((md (main-document document))
	   (root (opc:xml-root md)))
      (plump:get-elements-by-tag-name root "w:sectPr"))))
