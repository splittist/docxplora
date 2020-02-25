;;;; docxplora.asd

(asdf:defsystem #:docxplora
  :description "Primitive DOCX footling"
  :author "John Q. Splittist <splittist@splittist.com>"
  :license  "MIT"
  :version "0.0.1"
  :serial t
  :depends-on (#:alexandria #:uiop #:zip #:flexi-streams #:plump #:lquery)
  :components ((:file "package")
	       (:file "opc-content-types")
	       (:file "opc-relationship-types")
	       (:file "opc-namespaces")
               (:file "opc")
	       (:file "ooxml-content-types")
	       (:file "ooxml")
               (:file "docxplora")

	       (:file "changes")))
