;;;; opc-core-properties.lisp

(cl:in-package #:opc)

(defparameter *cp-namespaces*
  '(("cp" . "http://schemas.openxmlformats.org/package/2006/metadata/core-properties")
    ("dc" . "http://purl.org/dc/elements/1.1/")
    ("dcterms" . "http://purl.org/dc/terms/")
    ("dcmitype" . "http://purl.org/dc/dcmitype/")
    ("xsi" . "http://www.w3.org/2001/XMLSchema-instance")))
    
(defparameter *cp-contents*
   '("dc:title"
     "dc:subject"
     "dc:creator"
     "cp:keywords"
     "dc:description"
     "cp:lastModifiedBy"
     "cp:revision"
     ("dcterms:created" "xsi:type" "dcterms:W3CDTF")
     ("dcterms:modified" "xsi:type" "dcterms:W3CDTF")))