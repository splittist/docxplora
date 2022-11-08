;;;; settings.lisp

(in-package #:docxplora)

(defgeneric find-document-setting (thing tag-name)
  (:method ((document document) tag-name)
    (find-document-setting (document-settings document) tag-name))
  (:method ((document-settings document-settings) tag-name)
    (find-document-setting (opc:xml-root document-settings) tag-name))
  (:method ((root plump-dom:root) tag-name)
    (let ((settings (get-first-element-by-tag-name root "w:settings")))
      (find-child/tag settings tag-name))))


