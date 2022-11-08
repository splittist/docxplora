;;;; docxplora.asd

(asdf:defsystem #:docxplora
  :description "Primitive DOCX footling"
  :author "John Q. Splittist <splittist@splittist.com>"
  :license  "MIT"
  :version "0.0.1"
  :in-order-to ((asdf:test-op (asdf:test-op :docxplora-test)))
  :depends-on (#:alexandria #:uiop #:serapeum
			    			    
			    #:zip #:flexi-streams
			    #:plump #:lquery
			    #:cl-ppcre

			    #:trees
			    #:local-time
			    
			    #:wuss
			    #:recolor
			    #:imagesniff)
  :components ((:file "package")
	       (:file "utils")

	       (:module "opc"
		:depends-on ("package")
		:components ((:file "opc-content-types")
			     (:file "opc-core-properties")
			     (:file "opc-relationship-types")
			     (:file "opc-namespaces")
			     (:file "opc" :depends-on ("opc-content-types"
						       "opc-relationship-types"
						       "opc-namespaces"
						       "opc-core-properties"))))

	       (:module "ooxml"
		:depends-on ("package" "opc")
		:components ((:file "ooxml-content-types")
			     (:file "ooxml")))

	       (:module "wml"
		:depends-on ("package" "opc" "ooxml")
		:components  ((:file "wml-ordering")
			      (:file "utils" :depends-on ("wml-ordering"))
			      (:file "wml-namespaces")
			      (:file "wml" :depends-on ("utils" "wml-namespaces"))
			      
			      (:file "settings" :depends-on ("wml"))
			      (:file "sections" :depends-on ( "wml"))
			      (:file "tables" :depends-on ("wml"))
			      (:file "styles" :depends-on ("wml"))
			      (:file "numbering" :depends-on ("wml"))
			      (:file "hyperlinks" :depends-on ("wml"))
			      (:file "images" :depends-on ("wml"))
			      (:file "comments" :depends-on ("wml"))
			      (:file "footnotes" :depends-on ("wml"))
			      (:file "endnotes" :depends-on ("wml"))))
	       (:module "sml"
		:depends-on ("package" "opc" "ooxml" "wml")
		:components ((:file "utils")
			     (:file "sml-namespaces")
			     (:file "future")
			     (:file "cell-table")
			     (:file "sml" :depends-on ("utils" "sml-namespaces"))
			     (:file "shared-strings" :depends-on ("sml"))
			     (:file "workbook-writer" :depends-on ("sml" "shared-strings"))
			     (:file "workbook-editor" :depends-on ("workbook-writer"))
			     (:file "worksheet" :depends-on ("sml" "shared-strings"))
			     (:file "table" :depends-on ("sml"))))
	       (:file "docxplora")))

(asdf:defsystem #:docxplora/test
  :depends-on (#:docxplora #:parachute)
  :components ((:file "build-tests"))
  :perform (asdf:test-op (op c) (uiop:symbol-call :parachute :test :test-package)))
