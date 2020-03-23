;;;; hyperlinks.lisp

(cl:in-package #:docxplora)

;;; Hyperlinks (external)

(defun find-hyperlink (source uri)
  (find uri (opc:get-relationships-by-type source (opc:rt "HYPERLINK"))
	:key #'opc:target-uri
	:test #'string-equal)) ; FIXME Genrally check string= / string-equal

(defun ensure-hyperlink (source uri)
  (alexandria:if-let ((existing (find-hyperlink source uri)))
    existing
    (opc:create-relationship source uri (opc:rt "HYPERLINK") :target-mode "External")))

