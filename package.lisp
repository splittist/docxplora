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
   #:flush-part
   #:ct
   #:ns
   #:rt
   #:get-relationships-by-type-code
   #:opc-xml-part
   #:parse
   #:serialize
   #:xml-root
   #:flush-part
   #:write-part
   #:ensure-xml
   #:create-xml-part))

(defpackage #:plump-utils
  (:use #:cl)
  (:export
   #:find-child/tag
   #:find-children/tag
   #:find-child/tag/val
   #:ensure-child/tag
   #:remove-child/tag
   #:remove-if/tag
   #:remove-if-not/tag
   #:make-element/attrs
   #:get-first-element-by-tag-name
   #:tagp
   #:find-ancestor-element
   #:preservep
   #:make-text-element
   #:increment-attribute))

(defpackage #:sml
  (:use #:cl #:plump-utils))

(defpackage #:docxplora
  (:use #:cl #:plump-utils)
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
   #:header
   #:footer

   #:add-main-document

   #:paragraphs-in-document-order
   
   #:make-inline-image

   #:add-style-definitions
   #:add-style
   #:remove-style
   #:find-style-by-id
   
   #:ensure-hyperlink

   #:add-numbering-definitions
   #:ensure-numbering-definitions
   #:make-numbering-start-override

   #:get-comment-by-id
   #:comment-id
   #:comment-author
   #:comment-initials
   #:comment-date
   #:comments-in-document-order

   #:sections
   #:paragraph-section
   
   #:get-footnote-by-id
   #:footnote-references-in-document-order
   #:footnotes-in-document-order
   #:footnote-reference-paragraph
   #:footnote-reference-section
   #:footnote-references-numbering
   
   #:get-endnote-by-id
   #:endnote-references-in-document-order
   #:endnotes-in-document-order
   #:endnote-reference-paragraph
   #:endnote-reference-section
   #:endnote-references-numbering
   
   #:find-child/tag
   #:find-children/tag
   #:find-child/tag/val
   #:ensure-child/tag
   #:remove-child/tag
   #:remove-if/tag
   #:remove-if-not/tag
   #:make-element/attrs
   #:get-first-element-by-tag-name
   #:tagp
   #:find-ancestor-element
   #:preservep
   #:make-text-element
   #:increment-attribute
   
   #:coalesce-adjacent-text
   ))


