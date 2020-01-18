(cl:in-package #:opc)

(defparameter *namespaces*
  '(("CONTENT_TYPES" "http://schemas.openxmlformats.org/package/2006/content-types")
    ("CORE_PROPERTIES" "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
    ("DIGITAL_SIGNATURES" "http://schemas.openxmlformats.org/package/2006/digital-signature")
    ("RELATIONSHIPS" "http://schemas.openxmlformats.org/package/2006/relationships")
    ("MARKUP_COMPATIBILITY" "http://schemas.openxmlformats.org/markup-compatibility/2006")))

(defun ns (name)
  (cadr (find name *namespaces* :key #'car :test #'string-equal)))
