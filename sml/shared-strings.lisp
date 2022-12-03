;;;; shared-strings.lisp

(in-package #:sml)

(defclass string-table ()
  ((%count
    :initform 0
    :accessor sst-count)
   (%dict
    :initform (make-hash-table :test #'equal)
    :accessor sst-dict)))

(defun make-string-table ()
  (make-instance 'string-table))

(defun string-index (string-table string)
  (let ((dict (sst-dict string-table)))
    (incf (sst-count string-table))
    (alexandria:ensure-gethash string dict (hash-table-count dict))))

(defun sst-unique-count (string-table)
  (hash-table-count (sst-dict string-table)))

(defun sst-replace-entry (string-table old new)
  (let ((dict (sst-dict string-table)))
    (alexandria:when-let ((index (gethash old dict)))
      (remhash old dict)
      (setf (gethash new dict) index))))
      


