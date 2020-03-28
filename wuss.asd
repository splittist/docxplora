;;;; wuss.asd

(asdf:defsystem #:wuss
  :description "Word-focused Unsatisfactory Style Sheets"
  :author "John Q. Splittist <splittist@splittist.com>"
  :license  "MIT"
  :version "0.0.1"
  :depends-on (#:alexandria #:uiop #:serapeum
			    #:split-sequence
			    #:plump)
  :components ((:file "wuss")))
