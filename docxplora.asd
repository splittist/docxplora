;;;; docxplora.asd

(asdf:defsystem #:docxplora
  :description "Primitive DOCX footling"
  :author "John Q. Splittist <splittist@splittist.com>"
  :license  "MIT"
  :version "0.0.1"
  :depends-on (#:alexandria #:uiop #:serapeum
			    #:zip #:flexi-streams
			    #:plump #:lquery

			    #:wuss
			    #:recolor
			    #:imagesniff)
  :components ((:file "package")

	       (:file "opc-content-types" :depends-on ("package"))
	       (:file "opc-core-properties" :depends-on ("package"))
	       (:file "opc-relationship-types" :depends-on ("package"))
	       (:file "opc-namespaces" :depends-on ("package"))
               (:file "opc" :depends-on ("package"
					 "opc-content-types"
					 "opc-relationship-types"
					 "opc-namespaces"
					 "opc-core-properties"))

	       (:file "ooxml-content-types" :depends-on ("package"))
	       (:file "ooxml" :depends-on ("package" "opc"))

	       (:file "wml-ordering" :depends-on ("package"))
	       (:file "utils" :depends-on ("package" "wml-ordering"))
	       (:file "wml-namespaces" :depends-on ("package"))
	       (:file "wml" :depends-on ("package" "ooxml" "utils" "wml-namespaces"))

	       (:file "styles" :depends-on ("package" "wml"))
	       (:file "hyperlinks" :depends-on ("package" "wml"))
	       (:file "images" :depends-on ("package" "wml"))

	       (:file "docxplora")))
