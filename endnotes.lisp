;;;; endnotes.lisp

(in-package #:docxplora)

(defun get-endnote-by-id (root id)
  (when (numberp id) (setf id (princ-to-string id)))
  (let ((endnotes (plump:get-elements-by-tag-name root "w:endnote")))
    (find-if (lambda (e) (equal id (plump:attribute e "w:id"))) endnotes)))
