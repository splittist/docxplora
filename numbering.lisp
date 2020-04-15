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

(defun find-child/tag/val (root tag)
  (alexandria:when-let ((element (find-child/tag root tag)))
    (plump:attribute element "w:val")))

(defgeneric find-num-by-id (target numid)
  (:method ((document document) (numid string))
    (alexandria:when-let ((numbering-definitions (numbering-definitions document)))
      (let ((nums (plump:get-elements-by-tag-name (opc:xml-root numbering-definitions)
						  "w:num")))
	(find-if (alexandria:curry #'string= numid)
		 nums
		 :key (alexandria:rcurry #'plump:attribute "w:numId"))))))

(defgeneric find-abstract-num-by-id (target id)
  (:method ((document document) (id string))
    (alexandria:when-let ((numbering-definitions (numbering-definitions document)))
      (let ((abstract-nums (plump:get-elements-by-tag-name (opc:xml-root numbering-definitions)
							   "w:abstractNum")))
	(find-if (alexandria:curry #'string= id)
		 abstract-nums
		 :key (alexandria:rcurry #'plump:attribute "w:abstractNumId"))))))

(defun find-override (num ilvl)
  (find-if (alexandria:curry #'string= ilvl)
	   (plump:get-elements-by-tag-name num "w:lvlOverride")
	   :key (alexandria:rcurry #'plump:attribute "w:ilvl")))

(defun find-lvl (abstract-num ilvl)
  (find-if (alexandria:curry #'string= ilvl)
	   (plump:get-elements-by-tag-name abstract-num "w:lvl")
	   :key (alexandria:rcurry #'plump:attribute "w:ilvl")))

(defun find-lvl-by-style-id (abstract-num style-id)
  (find-if (alexandria:curry #'string= style-id)
	   (plump:get-elements-by-tag-name abstract-num "w:lvl")
	   :key (alexandria:rcurry #'find-child/tag/val "w:pStyle")))

(defclass list-info ()
  ((%style-id :initarg :style-id :initform nil :reader style-id)
   (%style-num :initarg :style-num :initform nil :reader style-num)
   (%style-abstract-num :initarg :style-abstract-num :initform nil :reader style-abstract-num)
   (%style-ilvl :initarg :style-ilvl :initform nil :reader style-ilvl)
   (%para-num :initarg :para-num :initform nil :reader para-num)
   (%para-abstract-num :initarg :para-abstract-num :initform nil :reader para-abstract-num)
   (%para-ilvl :initarg :para-ilvl :initform nil :reader para-ilvl)
   (%para-override :initarg :para-override :initform nil :reader para-override)))

(defmethod print-object ((object list-info) stream)
  (print-unreadable-object (object stream :type t :identity t)
    (if (para-abstract-num object)
	(format stream "~A / ~A"
		(plump:attribute (para-abstract-num object) "w:abstractNumId")
		(plump:attribute (para-ilvl object) "w:ilvl"))
	(format stream "~A" (style-id object)))))

(defgeneric list-num (object)
  (:method ((object list-info))
    (or (para-num object)
	(style-num object))))

(defgeneric list-level (object)
  (:method ((object list-info))
    (with-accessors ((para-ilvl para-ilvl)
		     (style-ilvl style-ilvl))
	object
      (or (and para-ilvl (parse-integer (plump:attribute para-ilvl "w:ilvl")))
	  (and style-ilvl (parse-integer (plump:attribute style-ilvl "w:ilvl")))))))

(defgeneric list-start (object)
  (:method ((object list-info))
    (when (or (para-ilvl object) (style-ilvl object))
      (or
       (find-child/tag/val (or (para-override object)
			       (para-ilvl object)
			       (style-ilvl object))
			   "w:start")
       "0"))))

(defgeneric list-restart (object)
  (:method ((object list-info))
    (when (or (para-ilvl object) (style-ilvl object))
      (find-child/tag/val (or (para-override object)
			      (para-ilvl object)
			      (style-ilvl object))
			  "w:lvlRestart"))))

(defgeneric list-is-legal (object)
  (:method ((object list-info))
    (when (or (para-ilvl object) (style-ilvl object))
      (serapeum:true
       (find-child/tag (or (para-override object)
			   (para-ilvl object)
			   (style-ilvl object))
		       "w:isLgl")))))

(defgeneric list-text (object)
  (:method ((object list-info))
    (when (or (para-ilvl object) (style-ilvl object))
      (find-child/tag/val (or (para-override object)
			      (para-ilvl object)
			      (style-ilvl object))
			  "w:lvlText"))))

(defgeneric list-format (object)
  (:method ((object list-info))
    (when (or (para-ilvl object) (style-ilvl object))
      (or
       (find-child/tag/val (or (para-override object)
			       (para-ilvl object)
			       (style-ilvl object))
			   "w:numFmt")
       "decimal"))))

(defgeneric list-start-override (object)
  (:method ((object list-info))
    (when (para-override object)
      (parse-integer
       (find-child/tag/val (para-override object) "w:startOverride")))))

(defun make-paragraph-list-info (document paragraph)
  (let (sid sn san sl pn pan pl or)
    (serapeum:and-let*
	((ppr (find-child/tag paragraph "w:pPr"))
	 (style-id (find-child/tag/val ppr "w:pStyle"))
	 (style-definition (find-style-by-id document style-id)))
      (serapeum:and-let*
	  ((style-ppr (find-child/tag style-definition "w:pPr"))
	   (style-npr (find-child/tag style-ppr "w:numPr"))
	   (style-numid (find-child/tag/val style-npr "w:numId"))
	   (style-num (find-num-by-id document style-numid))
	   (style-abstract-num (find-abstract-num-by-id
				document
				(find-child/tag/val style-num "w:abstractNumId"))))
	(setf sid style-id
	      sn style-num
	      san style-abstract-num
	      sl (find-lvl-by-style-id style-abstract-num style-id)))
      (serapeum:and-let*       
	  ((numpr (find-child/tag ppr "w:numPr"))
	   (numid (find-child/tag/val numpr "w:numId"))
	   (ilvl (find-child/tag/val numpr "w:ilvl"))
	   (num (find-num-by-id document numid))
	   (abstract-num (find-abstract-num-by-id
			  document
			  (find-child/tag/val num "w:abstractNumId"))))
	(setf pn num
	      pan abstract-num
	      pl (find-lvl abstract-num ilvl)
	      or (find-override num ilvl))))
    (make-instance 'list-info
		   :style-id sid
		   :style-num sn
		   :style-abstract-num san
		   :style-ilvl sl
		   :para-num pn
		   :para-abstract-num pan
		   :para-ilvl pl
		   :para-override or)))

(defun format-number (format number)
  (funcall (second (assoc format *numbering-formats* :test #'string=)) number))

(defparameter *numbering-formats*
  `(("aiueo" ,(lambda (num)
	       (let ((num (1+ (mod num 46))))
		 (cond
		   ((<= 1 num 44)
		    (string (code-char (+ num #xFF70))))
		   ((= num 45)
		    (string (code-char #xFF66)))
		   ((= num 46)
		    (string (code-char #xFF9D)))
		   (t (error "huh?"))))))
    ("cardinalText" ,(lambda (num) ;; FIXME - lang
		      (format nil "~@(~R~)" num)))
    ("chicago" ,(let ((symbols (list #x002A #x2020 #x2021 #x00A7)))
		 (lambda (num)
		   (multiple-value-bind (remainder quotient)
		       (floor (1- num) 4)
		     (format nil "~V@{~A~:*~}"
			     (1+ remainder)
			     (code-char (nth quotient symbols)))))))
    ("decimal" ,(lambda (num)
		  (format nil "~D" num)))
    ("decimalZero" ,(lambda (num)
		      (format nil "~2,'0D" num)))
    ("hex" ,(lambda (num)
	      (format nil "~X" num)))
    ("lowerLetter" ,(lambda (num) ;; FIXME - lang
		      (multiple-value-bind (remainder quotient)
			  (floor (1- num) 26)
			(format nil "~V@{~A~:*~}"
				(1+ remainder)
				(code-char (+ quotient (char-code #\a)))))))
    ("lowerRoman" ,(lambda (num)
		     (format nil "~(~@R~)" num)))
    ("none" ,(constantly "")) ;; ??
    ("numberInDash" ,(lambda (num)
		       (format nil "-~D-" num)))
    ("ordinal" ,(lambda (num) ;; FIXME - lang
		  (let* ((cardinal (princ-to-string num))
			 (suffix
			  (cond
			    ((serapeum:string$= "11" cardinal) "th")
			    ((serapeum:string$= "12" cardinal) "th")
			    ((serapeum:string$= "1" cardinal)  "st")
			    ((serapeum:string$= "2" cardinal)  "nd")
			    ((serapeum:string$= "3" cardinal)  "rd")
			    (t                        "th"))))
		    (format nil "~A~A" cardinal suffix))))
    ("ordinalText" ,(lambda (num) ;; FIXME - lang
		      (format nil "~@(~:R~)" num)))
    ("upperLetter" ,(lambda (num) ;; FIXME - lang
		      (multiple-value-bind (remainder quotient)
			  (floor (1- num) 26)
			(format nil "~V@{~A~:*~}"
				(1+ remainder)
				(code-char (+ quotient (char-code #\A)))))))
    ("upperRoman" ,(lambda (num)
		     (format nil "~@R" num)))))

(defun previous-formats (list-info)
  (let ((level (list-level list-info))
	(abstract-num (or (para-abstract-num list-info)
			  (style-abstract-num list-info))))
    (loop for l from (1- level) downto 0
       collecting
	 (let ((lvl (find-lvl abstract-num (princ-to-string l))))
	   (find-child/tag/val lvl "w:numFmt")))))

(defun format-level-text (list-text list-formats counts)
  (with-output-to-string (s)
    (when list-text
      (loop with percent-seen = nil
	 for char across list-text
	 do (cond
	      ((and percent-seen (digit-char-p char))
	       (let ((level (1- (digit-char-p char))))
		 (write-string (format-number (nth level list-formats)
					      (elt counts level))
			       s)
		 (setf percent-seen nil)))
	      (percent-seen
	       (write-char #\% s)
	       (write-char char s)
	       (setf percent-seen nil))
	      ((char= #\% char)
	       (setf percent-seen t))
	      (t
	       (write-char char s)))))))
       
#+(or)(defmacro do-story-paragraphs ((paragraph story &optional return) &body body)
  `(dolist (,paragraph (paragraphs-in-document-order (opc:xml-root ,story)) ,return)
     ,@body))

(defun paragraph-formatted-numbers (list-infos)
  (let ((count-dict (make-hash-table)))
    (map 'vector
	 (lambda (li)
	   (let ((list-num (list-num li)))
	     (when list-num
	       (unless (gethash list-num count-dict)
		 (setf (gethash list-num count-dict) (make-array 9 :initial-element 0)))
	       (let ((level (list-level li)))
		 (alexandria:if-let ((start-override (list-start-override li)))
		   (setf (aref (gethash list-num count-dict) level) start-override)
		   (incf (aref (gethash list-num count-dict) level)))
		 (loop for l from (1+ level) below 9 ;; FIXME - start, lvlRestart
		    do (setf (aref (gethash list-num count-dict) l) 0))
		 (let ((list-format (list-format li)))
		   (if (string= "bullet" list-format)
		       (list-text li)
		       (format-level-text
			(list-text li)
			(append
			 (if (list-is-legal li)
			     (make-list level :initial-element "decimal")
			     (previous-formats li))
			 (list (list-format li)))
			(gethash list-num count-dict))))))))
	     list-infos)))

