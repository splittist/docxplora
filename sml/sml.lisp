;;; sml.lisp

(in-package #:sml)

(defclass sml-document (docxplora:document)
  ())

(defun open-document (pathname)
  (let ((opc-package (opc:open-package pathname)))
    (make-instance 'sml-document :package opc-package)))

(defun make-document ()
  (let ((opc-package (make-instance 'opc:opc-package)))
    (make-instance 'sml-document :package opc-package)))

(defclass sml-part (opc:opc-part) ())

(defclass sml-xml-part (opc:opc-xml-part) ())

(defun create-sml-xml-part (class-name package uri content-type)
  (let ((part (opc:create-xml-part package uri content-type)))
    (change-class part class-name)))

(defclass workbook (sml-xml-part)())

(defmethod workbook ((document sml-document))
  (let* ((package (docxplora:opc-package document))
         (rel (first (opc:get-relationships-by-type-code package "OFFICE_DOCUMENT")))
         (target (opc:uri-merge "/" (opc:target-uri rel))))
    (docxplora:get-part-by-name document target :xml t :class 'workbook)))

(defgeneric add-workbook (document &key strict)
  (:method ((document sml-document) &key strict)
    (let* ((package (docxplora:opc-package document))
           (part (create-sml-xml-part
                  'workbook
                  package
                  "/xl/workbook.xml"
                  (opc:ct "SML_SHEET_MAIN")))
           (root (opc:xml-root part))
           (wb (plump:make-element root "workbook")))
      (plump:make-element wb "sheets")
      (dolist (ns *workbook-namespaces*)
        (setf (plump:attribute wb (format nil "xmlns~@[:~A~]" (car ns))) (cdr ns)))
      (dolist (mc *workbook-ignorables*)
        (setf (plump:attribute wb (car mc)) (cdr mc)))
      (opc:create-relationship package "/xl/workbook.xml" (opc:rt "OFFICE_DOCUMENT" :strict strict))
      part)))

(defun ensure-workbook (document &key strict)
  (alexandria:if-let ((existing (workbook document)))
    existing
    (add-workbook document :strict strict)))

(defun workbook-target (document rt class)
  (let* ((workbook (workbook document))
         (rel (first (opc:get-relationships-by-type workbook rt))))
    (when rel
      (let ((target (opc:uri-merge (opc:part-name workbook) (opc:target-uri rel))))
        (docxplora:get-part-by-name document target :xml t :class class)))))

(defmacro define-xml-part (class uri content-type root-element namespaces ignorables relationship-type)
  (let ((add-name (alexandria:symbolicate 'add- class))
        (ensure-name (alexandria:symbolicate 'ensure- class))
        (pred-name (alexandria:symbolicate class (if (find #\- (symbol-name class) :test #'char=) '-p 'p))))
    `(progn
       (defclass ,class (sml-xml-part)())
       (defgeneric ,class (document)
         (:method ((document sml-document))
           (multiple-value-bind (transitional strict)
               (opc:rt ,relationship-type)
             (or (workbook-target document transitional ',class)
                 (workbook-target document strict ',class)))))
       (defgeneric ,add-name (document &key strict)
         (:method ((document sml-document) &key strict)
           (let* ((package (docxplora:opc-package document))
                  (wb (workbook document))
                  (part (create-sml-xml-part
                         ',class
                         package
                         ,uri
                         (opc:ct ,content-type)))
                  (xml-root (opc:xml-root part))
                  (root (plump:make-element xml-root ,root-element)))
             (dolist (nss ',namespaces)
               (let ((ns (assoc nss *workbook-namespaces* :test #'equal)))
                 (setf (plump:attribute root (format nil "xmlns~@[:~A~]" (car ns))) (cdr ns))))
             (when ,ignorables
	       (setf (plump:attribute root "mc:Ignorable") ,ignorables))
             (opc:create-relationship wb (opc:uri-relative "/xl/workbook.xml" ,uri)
                                      (opc:rt ,relationship-type :strict strict))
             part)))
       (defgeneric ,ensure-name (document)
         (:method ((document sml-document))
           (alexandria:if-let ((result (,class document)))
             result
             (,add-name document))))
       (defgeneric ,pred-name (thing)
         (:method (thing) nil)
         (:method ((thing ,class)) t)))))

;; 12.3.1 calculation-chain workbook calcChain
(define-xml-part calculation-chain
  "/xl/calcChain.xml"
  "SML_CALC_CHAIN"
  "calcChain"
  (nil)
  nil
  "CALC_CHAIN")

;; 12.3.4 connections workbook connections
(define-xml-part connections
  "/xl/connections.xml"
  "SML_CONNECTIONS"
  "connections"
  (nil) ; CHECK
  nil
  "CONNECTIONS")

;; 12.3.9 external-workbook-references workbook externalLink

;; 12.3.15 shared-string-table workbook sst
(define-xml-part shared-string-table
  "/xl/sharedStrings.xml"
  "SML_SHARED_STRINGS"
  "sst"
  (nil)
  nil
  "SHARED_STRINGS")

;; 12.3.20 styles workbook styleSheet
(define-xml-part styles
  "/xl/styles.xml"
  "SML_STYLES"
  "styleSheet"
  (nil "mc" "x14ac" "x16r2" "xr")
  "x14ac x16r2 xr"
  "STYLES")

(define-xml-part metadata
  "/xl/metadata.xml"
  "SML_SHEET_METADATA"
  "metadata"
  (nil "xlrd" "xda")
  nil
  "SHEET_METADATA")
