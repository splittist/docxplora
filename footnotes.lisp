;;;; footnotes.lisp

(in-package #:docxplora)

(defun get-footnote-by-id (root id)
  (when (numberp id) (setf i (princ-to-string id)))
  (let ((footnotes (plump:get-elements-by-tag-name root "w:footnote")))
    (find-if (lambda (f) (equal id (plump:attribute f "w:id"))) footnotes)))
