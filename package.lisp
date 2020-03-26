;;;; package.lisp

(in-package #:cl)

(defpackage #:opc
  (:nicknames #:com.splittist.opc)
  (:use #:cl)
  (:export
   #:opc-package
   #:make-package
   #:open-package
   #:flush-package
   #:save-package
   #:package-pathname
   #:opc-part
   #:part-name
   #:content-type
   #:content
   #:create-part
   #:delete-part
   #:get-parts
   #:get-part
   #:opc-relationship
   #:create-relationship
   #:delete-relationship
   #:get-relationships
   #:get-relationship
   #:get-relationships-by-type
   #:source-uri
   #:target-uri
   #:target-mode
   #:relationship-type
   #:relationship-id
   #:uri-extension
   #:uri-filename
   #:uri-directory
   #:uri-relative
   #:uri-merge
   #:opc-xml-part
   #:flush-part
   #:ct
   #:ns
   #:rt
   #:opc-xml-part
   #:xml-root
   #:flush-part
   #:write-part
   #:ensure-xml
   #:create-xml-part))

(defpackage #:docxplora
  (:use #:cl)
  (:export
   #:document
   #:opc-package
   #:open-document
   #:make-document
   #:save-document
   #:document-type
   #:get-part-by-name
   #:main-document

   #:comments
   #:document-settings
   #:endnotes
   #:footnotes
   #:font-table
   #:glossary-document
   #:numbering-definitions
   #:style-definitions
   #:web-settings
   #:headers
   #:footers

   #:add-main-document
   
   #:make-inline-image

   #:add-style-definitions
   #:add-style
   #:remove-style
   #:find-style-by-id
   
   #:create-hyperlink
   #:ensure-hyperlink

   #:add-numbering-definitions
   #:ensure-numbering-definitions
   #:make-numbering-start-override

   #:find-child/tag
   #:find-children/tag
   #:ensure-child/tag
   #:remove-child/tag
   #:make-element/attrs
   #:get-first-element-by-tag-name
   #:coalesce-adjacent-text
   ))


