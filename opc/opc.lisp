;;;; opc.lisp

(in-package #:opc)

(defclass opc-package ()
  ((%package-parts :initarg :parts :accessor package-parts)
   (%ctm :initarg :ctm :accessor content-type-map)
   (%package-relationships :initarg :relationships :accessor package-relationships)
   (%pathname :initarg :pathname :initform nil :accessor package-pathname))
  (:default-initargs
   :parts (make-hash-table :test 'equal)
    :relationships (make-hash-table :test 'equal)
    :ctm (make-ctm)))

(defmethod print-object ((object opc-package) stream)
  (print-unreadable-object (object stream :type t :identity t)
    (format stream "~A" (package-pathname object))))

(defun make-opc-package ()
  (make-instance 'opc-package))

(define-condition opc-error (type-error)
  ((part :initarg :part :reader part)
   (path :initarg :path :reader path))
  (:report (lambda (condition stream)
	     (format stream "No ~A part found in ~A."
		     (part condition)
		     (path condition)))))

(defclass opc-part ()
  ((%package :initarg :package :accessor opc-package)
   (%content :initarg :content :accessor content)
   (%name :initarg :name :accessor part-name)
   (%content-type :initarg :type :accessor content-type)
   (%part-relationships :initarg :relationships :accessor part-relationships))
  (:default-initargs :content nil :relationships (make-hash-table :test 'equal)))

(defmethod print-object ((object opc-part) stream)
  (print-unreadable-object (object stream :type t :identity t)
    (format stream "~A" (part-name object))))

(defun open-package (pathname)
  (let ((package (make-instance 'opc-package :pathname pathname))
	(parts (make-hash-table :test 'equal)))
    ;; extract parts
    (handler-case
	(zip:with-zipfile (zip pathname)
	  (zip:do-zipfile-entries (name entry zip)
	    (let* ((name (concatenate 'string "/" name))
		   (part (make-instance 'opc-part
					:package package
					:name name
					:content (zip:zipfile-entry-contents entry))))
	      (setf (gethash name parts) part))))
      (simple-error (condition)
	(declare (ignore condition))
	(error (make-condition 'type-error :expected-type "OPC package" :datum pathname))))
    ;; extract relationships
    (unless (gethash "/_rels/.rels" parts)
      (make-condition 'opc-error :part "/_rels/.rels" :path pathname))
    (loop for part being the hash-values of parts
       when (uri-rels-p (part-name part))
       do (let ((source-name (uri-rels-source (part-name part)))
		(rels (rels-part->rels part)))
	    (if (string= "/" source-name)
		(setf (package-relationships package) rels)
		(setf (part-relationships (gethash source-name parts)) rels))))
    ;; extract content-types
    (let ((ct-part (gethash "/[Content_Types].xml" parts)))
      (unless ct-part
	(make-condition 'opc-error :part "[Content_Types]" :path pathname))
      (setf (content-type-map package) (ct-part->ctm ct-part)
	    (package-parts package) parts))
    ;; set parts' content-types
    (maphash #'(lambda (name part) (setf (content-type part) (get-ct-from-ctm package name)))
	     parts)
    package))

(defun get-ct-from-ctm (package name)
  (let* ((ctm (content-type-map package))
	 (defaults (ctm-defaults ctm))
	 (overrides (ctm-overrides ctm))
	 (extension (uri-extension name)))
    (or (gethash name overrides)
	(gethash extension defaults))))

(defun list-parts (package)
  (format t "~{~A~%~}" (mapcar #'part-name (get-parts package))))

(defclass content-type-map ()
  ((%defaults :initarg :defaults :accessor ctm-defaults)
   (%overrides :initarg :overrides :accessor ctm-overrides)))

(defun make-ctm ()
  (let ((defaults (make-hash-table :test 'equal)))
    (setf (gethash "rels" defaults) (ct "OPC_RELATIONSHIPS")
	  (gethash "xml" defaults) (ct "XML"))
    (make-instance 'content-type-map
		   :defaults defaults
		   :overrides (make-hash-table :test 'equal))))

(defun parse (input &key root)
  (let ((plump:*tag-dispatchers* plump:*xml-tags*))
    (plump:parse input :root root)))

(defun serialize (node &optional (stream t))
  (let ((plump:*tag-dispatchers* plump:*xml-tags*))
    (plump:serialize node stream)))

(defun ct-part->ctm (ct-part)
  (let ((root (parse (flexi-streams:octets-to-string (content ct-part) :external-format :utf8)))
	(defaults (make-hash-table :test 'equal))
	(overrides (make-hash-table :test 'equal)))
    (loop for default in (plump:get-elements-by-tag-name root "Default")
       do (setf (gethash (plump:attribute default "Extension") defaults) (plump:attribute default "ContentType")))
    (loop for override in (plump:get-elements-by-tag-name root "Override")
       do (setf (gethash (plump:attribute override "PartName") overrides) (plump:attribute override "ContentType")))
    (make-instance 'content-type-map :defaults defaults :overrides overrides)))

(defun update-ctm (package)
  (let* ((ctm (content-type-map package))
	 (defaults (ctm-defaults ctm))
	 (overrides (ctm-overrides ctm))
	 (parts (get-parts package)))
    (dolist (part parts)
      (let* ((name (part-name part))
	     (extension (uri-extension name))
	     (ct (content-type part)))
	(cond
	  ((string= "/[Content_Types].xml" name)
	   nil)
	  ((not (member extension (alexandria:hash-table-keys defaults) :test #'string-equal))
	   (setf (gethash extension defaults) ct))
	  ((string= ct (gethash extension defaults))
	   nil)
	  (t
	   (setf (gethash name overrides) ct)))))))

(defun list-content-types (package)
  (let* ((ctm (content-type-map package))
	 (defaults (ctm-defaults ctm))
	 (overrides (ctm-overrides ctm)))
    (maphash #'(lambda (name type)(format t "Default ~A ~A~%" name type)) defaults)
    (maphash #'(lambda (name type)(format t "Override ~A ~A~%" name type)) overrides)))

(defclass opc-relationship ()
  ((%package :initarg :package :accessor opc-package)
   (%source-uri :initarg :source :accessor source-uri)
   (%target-uri :initarg :target :accessor target-uri)
   (%target-mode :initarg :mode :accessor target-mode)
   (%type :initarg :type :accessor relationship-type)
   (%id :initarg :id :accessor relationship-id)))

(defmethod print-object ((object opc-relationship) stream)
  (print-unreadable-object (object stream :type t :identity t)
    (format stream "~A" (relationship-id object))))

(defun rels-part->rels (rels-part)
  (let ((source (uri-rels-source (part-name rels-part)))
	(root (parse (flexi-streams:octets-to-string (content rels-part) :external-format :utf8)))
	(rels (make-hash-table :test 'equal)))
    (loop for rel in (plump:get-elements-by-tag-name root "Relationship")
       do (setf (gethash (plump:attribute rel "Id") rels)
		(make-instance 'opc-relationship
			       :package (opc-package rels-part)
			       :source source
			       :target (plump:attribute rel "Target")
			       :mode (or (plump:attribute rel "TargetMode") "Internal")
			       :type (plump:attribute rel "Type")
			       :id (plump:attribute rel "Id"))))
    rels))

(defgeneric create-part (package uri content-type) ; content?
  (:method ((package opc-package)(uri string)(content-type string))
    (let ((part (make-instance 'opc-part
			       :package package
			       :name uri
			       :type content-type)))
      (setf (gethash uri (package-parts package)) part)
      part)))

(defgeneric delete-part (package uri)
  (:method ((package opc-package)(uri string))
    (alexandria:when-let ((part (get-part package uri)))
      (setf (opc-package part) nil)
      (remhash uri (package-parts package)))))

(defgeneric get-parts (package)
  (:method ((package opc-package))
    (alexandria:hash-table-values (package-parts package))))

(defgeneric get-part (package uri)
  (:method ((package opc-package)(uri string))
    (gethash uri (package-parts package))))

(defun next-rid (relhash)
  (let ((ids (alexandria:hash-table-keys relhash)))
    (loop for i from 1
       for rid = (format nil "rId~D" i) then (format nil "rId~D" i)
       while (find rid ids :test 'string-equal)
	 finally (return rid))))
 
(defgeneric create-relationship (source uri relationship-type &key id target-mode)
  (:method ((source opc-package) (uri string) (relationship-type string) &key id (target-mode "Internal"))
    (let* ((rels (package-relationships source))
	   (id (or id (next-rid rels))))
      (setf (gethash id rels)
	    (make-instance 'opc-relationship
			   :id id
			   :type relationship-type
			   :mode target-mode
			   :source "/"
			   :package source
			   :target uri)))) ;; FIXME uri-relative
  (:method ((source opc-part) (uri string) (relationship-type string) &key id (target-mode "Internal"))
    (let* ((rels (part-relationships source))
	   (id (or id (next-rid rels))))
      (setf (gethash id rels)
	    (make-instance 'opc-relationship
			   :id id
			   :type relationship-type
			   :mode target-mode
			   :source (part-name source)
			   :package (opc-package source)
			   :target uri))))) ;; FIXME uri-relative

(defgeneric delete-relationship (source id)
  (:method ((source opc-package) (id string))
    (remhash id (package-relationships source)))
  (:method ((source opc-part) (id string))
    (remhash id (part-relationships source))))

(defgeneric get-relationships (source)
  (:method ((source opc-package))
    (alexandria:hash-table-values (package-relationships source)))
  (:method ((source opc-part))
    (alexandria:hash-table-values (part-relationships source))))

(defgeneric get-relationship (source id)
  (:method ((source opc-package)(id string))
    (gethash id (package-relationships source)))
  (:method ((source opc-part)(id string))
    (gethash id (part-relationships source))))

(defgeneric get-relationships-by-type (source type)
  (:method ((source opc-package)(type string))
    (remove-if-not #'(lambda (item-type) (string= item-type type))
		   (alexandria:hash-table-values (package-relationships source))
		   :key #'relationship-type))
  (:method ((source opc-part)(type string))
    (remove-if-not #'(lambda (item-type) (string= item-type type))
		   (alexandria:hash-table-values (part-relationships source))
		   :key #'relationship-type)))

;;;; URIs

;;; empty string is package?? turns out absolute uris start with "/"

(defun uri-extension (uri)
  (let ((dot (position #\. uri :from-end t)))
    (if dot
	(subseq uri (1+ dot))
	nil)))

(defun uri-filename (uri)
  (let ((slash (position #\/ uri :from-end t)))
    (if slash
	(subseq uri (1+ slash))
	uri)))

(defun uri-directory (uri)
  (let ((slash (position #\/ uri :from-end t)))
    (if slash
	(subseq uri 0 (1+ slash))
	""))) ;; or nil?

(defun uri-merge (source target)
  (namestring (uiop:parse-unix-namestring (namestring (merge-pathnames target source)))))

(defun uri-relative (source target)
  (enough-namestring target source))

(defun uri-rels-source (uri)
  (let ((preidx (search "_rels" uri :from-end t))
	(dot (position #\. uri :from-end t)))
    (concatenate 'string (subseq uri 0 preidx) (subseq uri (+ preidx 6) dot))))

(defun uri-source-rels (uri)
  (let ((slash (position #\/ uri :from-end t)))
    (if slash
	(concatenate 'string (subseq uri 0 slash) "/_rels" (subseq uri slash) ".rels")
	(concatenate 'string "_rels/" uri ".rels"))))

(defun uri-rels-p (uri)
  (and (search "_rels" uri)
       (string= "rels" (uri-extension uri))))

;;;; xml output

(defun make-opc-xml-header (parent)
  (let ((header (plump:make-xml-header parent)))
    (setf (plump:attribute header "version") "1.0"
	  (plump:attribute header "encoding") "UTF-8"
	  (plump:attribute header "standalone") "yes")
    header))

(defun make-content_types (package)
  (let ((root (plump:make-root)))
    (make-opc-xml-header root)
    (let ((types (plump:make-element root "Types")))
      (setf (plump:attribute types "xmlns")(ns "CONTENT_TYPES"))
      (let* ((ctm (content-type-map package))
	     (defaults (ctm-defaults ctm))
	     (overrides (ctm-overrides ctm)))
	(maphash #'(lambda (extension ct)
		     (let ((default (plump:make-element types "Default")))
		       (setf (plump:attribute default "Extension") extension
			     (plump:attribute default "ContentType") ct)))
		 defaults)
	(maphash #'(lambda (partname ct)
		     (let ((override (plump:make-element types "Override")))
		       (setf (plump:attribute override "PartName") partname
			     (plump:attribute override "ContentType") ct)))
		 overrides)))
    root))

(defun make-rels (rels-part)
  (let ((root (plump:make-root)))
    (make-opc-xml-header root)
    (let ((rels (plump:make-element root "Relationships")))
      (setf (plump:attribute rels "xmlns")(ns "RELATIONSHIPS"))
      (maphash #'(lambda (id opc-rel)
		   (let ((rel (plump:make-element rels "Relationship")))
		     (setf (plump:attribute rel "Id") id
			   (plump:attribute rel "Type") (relationship-type opc-rel)
			   (plump:attribute rel "Target") (target-uri opc-rel))
		     (if (string= "External" (target-mode opc-rel))
			 (setf (plump:attribute rel "TargetMode") "External"))))
	       rels-part))
    root))

(defun ensure-rels-part (package name)
  (unless (get-part package name)
    (create-part package name (ct "OPC_RELATIONSHIPS"))))

(defun ensure-content_types (package)
  (unless (get-part package "/[Content_Types].xml")
    (create-part package "/[Content_Types].xml" "")))

(defun user-parts (package)
  (remove-if
   #'(lambda (part)
       (let ((name (part-name part)))
	 (or (string= "/[Content_Types].xml" name)
	     (uri-rels-p name))))
   (get-parts package)))

(defun save-package (package &optional pathname)
  (let ((pathname (or pathname (package-pathname package))))
    (unless pathname (error "No pathname to save to"))
    (flush-package package)
    (write-zipfile package pathname)))

(defun flush-package (package)
  (generate-part-rels package)
  (generate-package-rels package)
  (generate-content_types package)
  (dolist (part (user-parts package))
    (flush-part part)))

(defun generate-part-rels (package)
  (let ((user-parts (user-parts package)))
    (dolist (part user-parts)
      (when (get-relationships part)
	(let ((rels-name (uri-source-rels (part-name part))))
	  (ensure-rels-part package rels-name)
	  (let ((rels-part (get-part package rels-name)))
	    (setf (content rels-part)
		  (xml-octets (make-rels (part-relationships part))))))))))

(defun generate-package-rels (package)
  (ensure-rels-part package (uri-source-rels "/"))
  (setf (content (get-part package (uri-source-rels "/")))
	(xml-octets (make-rels (package-relationships package)))))

(defun generate-content_types (package)
  (ensure-content_types package)
  (reset-ctm package)
  (update-ctm package)
  (setf (content (get-part package "/[Content_Types].xml"))
	(xml-octets (make-content_types package))))

(defun reset-ctm (package)
  (let* ((ctm (content-type-map package))
	 (defaults (ctm-defaults ctm))
	 (overrides (ctm-overrides ctm)))
    (clrhash overrides)
    (dolist (key (alexandria:hash-table-keys defaults))
      (unless (member key '("xml" "rels") :test #'string=)
	(remhash key defaults)))))

(defun write-zipfile (package pathname)
  (zip:with-output-to-zipfile (zip pathname :if-exists :supersede)
    (maphash #'(lambda (name part)
		 (let ((name (subseq name 1))) ;; strip leading "/"
		   (flexi-streams:with-input-from-sequence (content (content part))
		     (zip:write-zipentry zip name content :file-write-date (get-universal-time)))))
	     (package-parts package))))

(defun xml-octets (xml)
  (flexi-streams:string-to-octets (serialize xml nil) :external-format :utf8))

(defclass xml-root-mixin ()
  ((%xml-root :initarg :xml-root :accessor xml-root)))

(defclass opc-xml-part (opc-part xml-root-mixin) ())

(defun ensure-xml (part)
  (when (eql 'opc-part (type-of part))
    (change-class part 'opc-xml-part)
    (setf (xml-root part)
	  (parse (flexi-streams:octets-to-string (content part) :external-format :utf8))))
  part)

(defun create-xml-part (package uri content-type)
  (let ((part (create-part package uri content-type)))
    (ensure-xml part)
    (make-opc-xml-header (xml-root part))
    part))

(defgeneric flush-part (part)
  (:method (part)
    nil)
  (:method ((part opc-xml-part))
    (setf (content part)
	  (flexi-streams:string-to-octets
	   (serialize (xml-root part) nil)
	   :external-format :utf8))))

(defun write-part (part filespec &key (if-exists :supersede))
  (with-open-file (out filespec :direction :output :if-exists if-exists :element-type '(unsigned-byte 8) :external-format :utf8)
    (flush-part part)
    (write-sequence (content part) out)))

(defun add-core-properties-part (package)
  (let* ((part (create-xml-part package "/docProps/core.xml" (ct "OPC_CORE_PROPERTIES")))
	 (root (xml-root part))
	 (cp (plump:make-element root "cp:coreProperties")))
    (dolist (ns *cp-namespaces*)
      (setf (plump:attribute cp (format nil "xmlns:~A" (car ns))) (cdr ns)))
#+(or)    (dolist (entry *cp-contents*)
      (etypecase entry
	(string (plump:make-fulltext-element cp entry))
	(list (let ((element (plump:make-fulltext-element cp (first entry))))
		(setf (plump:attribute element (second entry)) (third entry))))))
    (create-relationship package "/docProps/core.xml" (rt "CORE_PROPERTIES"))
    part))
