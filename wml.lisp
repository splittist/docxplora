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

;;;;;;;;;;;;;;;;; building

(defun preservep (string) ; FIXME - stop defining this everywhere
  (or (alexandria:starts-with #\Space string :test #'char=)
      (alexandria:ends-with #\Space string :test #'char=)))

(defun make-text-element (parent text)
  (let ((wt (plump:make-element parent "w:t")))
    (when (preservep text)
      (setf (plump:attribute wt "xml:space") "preserve"))
    (plump:make-text-node wt text)
    wt))

(defun split-text-element (node index)
  (let* ((text (plump:text node))
	 (limit (length text)))
    (assert (<= 0 index limit) (index) "index out of range: ~D" index)
    (when (zerop index)
      (return-from split-text-element (values nil node)))
    (when (= index limit)
      (return-from split-text-element (values node nil)))
    (let* ((left (subseq text 0 index))
	   (left-space (preservep left))
	   (right (subseq text index)))
      (setf (plump:text (plump:first-child node)) left)
      (if left-space
	  (setf (plump:attribute node "xml:space") "preserve")
	  (plump:remove-attribute node "xml:space"))
      (let ((new (make-text-element (plump:make-root) right)))
	(plump:insert-after node new)
	(values node new)))))

(defun remove-if/tag (tag sequence)
  (plump:ensure-child-array
   (remove-if (alexandria:curry #'string= tag)
	      sequence
	      :key #'plump:tag-name)))

(defun remove-if-not/tag (tag sequence)
  (plump:ensure-child-array
   (remove-if-not (alexandria:curry #'string= tag)
		  sequence
		  :key #'plump:tag-name)))

(defun split-run (run index)
  (cond ((zerop index) ; the beginning
	 (values nil run))
	((= index (text-length run)) ; the end
	 (values run nil))
	(t
	 (multiple-value-bind (child sub-index)
	     (child-at-index run index)
	   (let* ((split-text (not (zerop sub-index)))
		  (child-position (plump:child-position child))
		  (children (plump:children run))
		  (left-children
		   (remove-if/tag "w:rPr" (subseq children 0 child-position)))
		  (right-children
		   (remove-if/tag "w:rPr" (subseq children (+ child-position
							      (if split-text 1 0)))))
		  (left (plump:make-element (plump:make-root) "w:r"
					    :children left-children))
		  (right (plump:make-element (plump:make-root) "w:r"
					     :children right-children)))
	     (when split-text
	       (multiple-value-bind (left-text right-text)
		   (split-text-element child sub-index)
		 (plump:append-child left left-text)
		 (plump:prepend-child right right-text)))
	     (alexandria:when-let ((rpr (clone-run-properties run)))
	       (plump:prepend-child left rpr)
	       (plump:prepend-child right rpr))
	     (values left right))))))

;; FIXME - make these data-driven
;; FIXME - w:ins / w:del and other paragraph children

(defun text-length (node &optional (start 0))
  (+ start
     (if (typep node 'plump-dom:element)
	 (serapeum:string-case (plump:tag-name node)
	   ("w:br" 1) ; FIXME - all breaks?
	   ("w:cr" 1)
	   (("w:delInstrText" "w:instrText") 0) ; FIXME - ??
	   (("w:delText" "w:t") (length (plump:text node)))
	   ("w:noBreakHyphen" 1)
	   ;;("w:ptab" ?)
	   ("w:sym" 1)
	   ("w:tab" 1)
       
	   ("w:r" (loop for child across (plump:children node)
		     summing (text-length child start)))
	   ("w:p" (loop for child across (plump:children node)
		     summing (text-length child start)))
	   
	   (t 0))
	 0)))

(defun to-text (node)
  (let ((ac '()))
    (labels
	((acc (node)
	   (when (typep node 'plump-dom:element)
	     (serapeum:string-case (plump:tag-name node)
	       ("w:br" (push #.(string #\Newline) ac))
	       ("w:cr" (push #.(string #\Newline) ac))
	       (("w:delInstrText" "w:instrText") nil)
	       (("w:delText" "w:t") (push (plump:text node) ac))
	       ("w:noBreakHyhen" (push #.(string #\-) ac))
	       ;;("w:ptab" ?)
	       ("w:sym" (push #.(string #\*) ac)) ;; FIXME - ??
	       ("w:tab" (push #.(string #\Tab) ac))
	       
	       ("w:r" (loop for child across (plump:children node)
			 do (acc child)))
	       ("w:p" (loop for child across (plump:children node)
			 do (acc child))
		      #+(or)(push #.(string #\Newline) ac)))))) ; no #\Newline
      (acc node))
    (serapeum:string-join (nreverse ac))))

(defparameter *run-escape-table* ;; FIXME - don't keep defining this
  (serapeum:dictq
   #\< "&lt;"
   #\> "&gt;"
   #\" "&quot;"
   #\' "&apos;"))

(defun run-from-text-preprocess (string)
  (let ((result '())
	(text '())
	(escaped (serapeum:escape string *run-escape-table*)))
    (loop for char across escaped
       do (case char
	    (#\Tab (when text
		     (push (coerce (nreverse text) 'string) result)
		     (setf text nil))
		   (push :tab result))
	    (#\Newline (when text
			 (push (coerce (nreverse text) 'string) result)
			 (setf text nil))
		       (push :cr result))
	    (t (push char text)))
       finally (progn (when text
			(push (coerce (nreverse text) 'string) result))
		      (return (nreverse result))))))

(defun run-from-text (string &optional run-properties)
  (let ((prelist (run-from-text-preprocess string))
	(run (plump:make-element (plump:make-root) "w:r")))
    (when run-properties
      (plump:prepend-child run run-properties))
    (dolist (item prelist run)
      (cond
	((eq :tab item)
	 (plump:make-element run "w:tab"))
	((eq :cr item)
	 (plump:make-element run "w:cr"))
	((stringp item)
	 (make-text-element run item))
	(t (error "Unknown item"))))))

(defun paragraphs-in-document-order (node)
  (let ((result '()))
    (plump:traverse
     node
     #'(lambda (node) (when (tagp node "w:p") (push node result)))
     :test #'plump:element-p)
    (nreverse result)))

(defun tables-in-document-order (node)
  (let ((result '()))
    (plump:traverse
     node
     #'(lambda (node) (when (tagp node "w:tbl") (push node result)))
     :test #'plump:element-p)
    (nreverse result)))

(defun child-at-index (node index)
  (assert (<= 0 index (text-length node)))
  (loop with i = 0
     with prev-i = 0
     until (> i index)
     for child across (plump:children node)
     do (shiftf prev-i i (+ i (text-length child)))
     finally (return
	       (if (> index i)
		 nil
		 (values child (- index prev-i))))))

(defun clone-run-properties (run)
  (check-type run plump-dom:element)
  (assert (tagp run "w:r"))
  (alexandria:when-let ((rpr (find-child/tag run "w:rPr")))
    (plump:clone-node rpr t)))

(defun clone-paragraph-properties (paragraph)
  (check-type paragraph plump-dom:element)
  (assert (tagp paragraph "w:p"))
  (alexandria:when-let ((ppr (find-child/tag paragraph "w:pPr")))
    (plump:clone-node ppr t)))

(defun merge-properties (destination source)
  (let ((dchildren (plump:children destination))
	(schildren (plump:children source)))
    (serapeum:do-each (schild schildren destination)
      (unless (find (plump:tag-name schild) dchildren :key #'plump:tag-name :test #'string=)
	(plump:append-child destination (plump:clone-node schild t))))))

(defun paragraph-append-text (paragraph string &optional run-properties)
  (let ((run (run-from-text string run-properties)))
    (plump:append-child paragraph run)
    paragraph))

(defun paragraph-insert-text (paragraph index text &optional run-properties)
  ;; let new-run run-from-text text run-properties
  ;; let limit text-length paragraph
  ;; when = index limit -> append paragraph new-run
  ;; when zerop index -> prepend [but pPr] paragraph new-run
  ;; let old-run, sub-idx child-at-index paragraph index
  ;; when/cond old-run -> left, right = split-run old-run idx
  ;; replace old-run with left, run, right
