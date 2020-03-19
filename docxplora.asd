;;;; docxplora.asd

(asdf:defsystem #:docxplora
  :description "Primitive DOCX footling"
  :author "John Q. Splittist <splittist@splittist.com>"
  :license  "MIT"
  :version "0.0.1"
  :serial t
  :depends-on (#:alexandria #:uiop #:serapeum
			    #:zip #:flexi-streams
			    #:plump #:lquery)
  :components ((:file "package")
	       (:file "wuss")
	       (:file "recolor")
	       (:file "opc-content-types")
	       (:file "opc-relationship-types")
	       (:file "opc-namespaces")
               (:file "opc")
	       (:file "ooxml-content-types")
	       (:file "wml-namespaces")
	       (:file "wml-ordering")
	       (:file "ooxml")
               (:file "docxplora")))
