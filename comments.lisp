;;;; comments.lisp

(in-package #:docxplora)

(defun get-comment-by-id (root id)
  (when (numberp id) (setf id (princ-to-string id)))
  (let ((comments (plump:get-elements-by-tag-name root "w:comment")))
    (find-if (lambda (c) (equal id (plump:attribute c "w:id"))) comments)))

(defun comment-id (comment)
  (plump:attribute comment "w:id"))

(defun comment-author (comment)
  (plump:attribute comment "w:author"))

(defun comment-initials (comment)
  (plump:attribute comment "w:initials"))

(defun comment-date (comment)
  (plump:attribute comment "w:date"))


