;;;; imagesniff.asd

(asdf:defsystem #:imagesniff
  :description "Simple information about popular image types"
  :author "John Q. Splittist <splittist@splittist.com>"
  :license  "MIT"
  :version "0.0.1"
  :depends-on (#:binary-io #:babel-streams #:alexandria)
  :components ((:file "imagesniff")))
