;;;; recolor.asd

(asdf:defsystem #:recolor
  :description "Convert colorize output to wml"
  :author "John Q. Splittist <splittist@splittist.com>"
  :license  "MIT"
  :version "0.0.1"
  :depends-on (#:plump #:serapeum #:alexandria
		       #:split-sequence #:colorize)
  :components ((:file "recolor")))
