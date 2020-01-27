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


(defun find-child (parent child-tag-name)
  (find child-tag-name (plump:children parent) :key #'plump:tag-name :test #'equal))

(defun fix-character-style (run style-name &optional remove-color)
  (let* ((run-props (or (find-child run "w:rPr")
			(plump:make-element run "w:rPr")))
	 (run-style (or (find-child run "w:rStyle")
			(plump:make-element run-props "w:rStyle"))))
    (setf (plump:attribute run-style "w:val") style-name)
    (when remove-color
      (alexandria:when-let (color (find-child run-props "w:color"))
	(plump:remove-child color))))
  run)

(lquery:define-lquery-function fix-character-style (run style-name &optional remove-color)
  (fix-character-style run style-name remove-color))

(defun handle-insertions (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::p" "w::ins" "w::r"
	      (fix-character-style "IPCInsertion")
	      (remove-attr "w:rsidR" "w:rsidRPr")
	      (unwrap))))

(defun handle-move-tos (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::moveTo" "w::r"
	      (fix-character-style "IPCMoveTo")
	      (remove-attr "w:rsidR" "w:rsidRPr")
	      (unwrap))
    (lquery:$ "w::moveToRangeStart"
	      (add "w::moveToRangeEnd")
	      (remove))))

(defun handle-move-froms (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::moveFrom" "w::r"
	      (fix-character-style "IPCMoveFrom")
	      (remove-attr "w:rsidR" "w:rsidRPr" "w:rsidDel")
	      (unwrap))
    (lquery:$ "w::moveFromRangeStart"
	      (add "w::moveFromRangeEnd")
	      (remove))))

(defun fix-del-text (run)
  (loop for child across (plump:children run)
     when (equal "w:delText" (plump:tag-name child))
     do (setf (plump:tag-name child) "w:t"))
  run)

(lquery:define-lquery-function fix-del-text (run)
  (fix-del-text run))

(defun handle-deletions (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::p" "w::del" "w::r"
	      (fix-character-style "IPCDeletion")
	      (fix-del-text)
	      (remove-attr "w:rsidR" "w:rsidRPr" "w:rsidDel")
	      (unwrap))))

(defun process-document (document)
  (let ((root (opc:xml-root (main-document document))))
    (handle-insertions root)
    (handle-deletions root)
    (handle-move-tos root)
    (handle-move-froms root))
  document)

(defun make-element/attrs (root tag-name &rest attributes)
  (plump:make-element root tag-name :attributes (alexandria:plist-hash-table attributes :test 'equalp)))

(defun make-ipc-style (id name color &key underline strike dstrike)
  (when (and strike dstrike)
    (error "Can't have both single and double strikethrough."))
  (let* ((root (plump:make-root))
	 (style (make-element/attrs root "w:style" "w:type" "character" "w:customStyle" "1" "w:styleId" id)))
    (make-element/attrs style "w:name" "w:val" name)
    (make-element/attrs style "w:basedOn" "w:val" "DefaultParagraphFont")
    (make-element/attrs style "w:uiPriority" "w:val" "1")
    (make-element/attrs style "w:qFormat")
    (let ((run-props (plump:make-element style "w:rPr")))
      (make-element/attrs run-props "w:color" "w:val" color)
      (when underline
	(make-element/attrs run-props "w:u" "w:val" underline))
      (when strike
	(plump:make-element run-props "w:strike"))
      (when dstrike
	(plump:make-element run-props "w:dstrike")))
    style))
