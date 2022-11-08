;;;; sml/utils.lisp

(in-package #:sml)

(defparameter *r1c1-scanner* (cl-ppcre:create-scanner "R(\\d+)C(\\d+)"
						      :case-insensitive-mode nil
						      :single-line-mode t))

(defparameter *a1-scanner* (cl-ppcre:create-scanner "(\\$?)([A-Z]+)(\\$?)(\\d+)"
						    :case-insensitive-mode t
						    :single-line-mode t))

(defun char-digit (char)
  (1+ (- (char-code char) #.(char-code #\A))))

(defun a1-r1 (column-reference)
  (reduce (lambda (x y) (+ (* 26 x) (char-digit y)))
	  (coerce column-reference 'list)
	  :initial-value 0))

(defun not-empty-p (sequence)
  (not (alexandria:emptyp sequence)))

(defun decode-a1 (reference)
  (cl-ppcre:register-groups-bind ((#'not-empty-p abs-col)
				  (#'a1-r1 col)
				  (#'not-empty-p abs-row)
				  (#'parse-integer row))
      (*a1-scanner* (string-upcase (string reference)))
    (values row col abs-row abs-col)))

(defun decode-a1-range (reference)
  (destructuring-bind (start end)
      (serapeum:split-sequence #\: reference)
    (multiple-value-bind (start-row start-col)
	(decode-a1 start)
      (if end
	  (multiple-value-bind (end-row end-col)
	      (decode-a1 end)
	    (values start-row start-col end-row end-col))
	  (values start-row start-col start-row start-col)))))

(defun encode-a1 (row col)
  (format nil "~A~D" (r1-a1 col) row))

(defun encode-a1-range (start-row start-col end-row end-col)
  (if (and (= start-row end-row)
	   (= start-col end-col))
      (encode-a1 start-row start-col)
      (format nil "~A:~A" (encode-a1 start-row start-col) (encode-a1 end-row end-col))))

(defun a1-r1c1 (reference)
  (multiple-value-bind (col row abs-col abs-row)
      (decode-a1 reference)
    (format nil "~:[~;$~]R~D~:[~;$~]C~D" abs-col col abs-row row)))

(defun r1-a1 (column-reference)
  (let ((str '()))
    (loop while (plusp column-reference)
	  do (let ((remainder (mod column-reference 26)))
	       (when (zerop remainder) (setf remainder 26))
	       (push (code-char (+ #.(1- (char-code #\A)) remainder)) str)
	       (setf column-reference (floor (/ (1- column-reference) 26))))
	  finally (return (coerce str 'string)))))

(defun r1c1-a1 (reference)
  (cl-ppcre:register-groups-bind ((#'parse-integer row col))
      (*r1c1-scanner* reference)
    (encode-a1 row col)))

(defparameter *xstring-xcode-scanner* (cl-ppcre:create-scanner "(_x[0-9a-fA-F]{4}_)"))

(defparameter *xstring-control-chars-scanner* (cl-ppcre:create-scanner "[\\x00-\\x08\\x0b-\\x1f]"))

(defun string-xstring (string)
  (serapeum:~> string
	       (cl-ppcre:regex-replace-all *xstring-xcode-scanner* _ "_x005F\\1")
	       (cl-ppcre:regex-replace-all *xstring-control-chars-scanner*
					   _
					   (lambda (match) (format nil "_x~4,'0x_" (char-code (char match 0))))
					   :simple-calls t)
	       (serapeum:string-replace #.(string #\UFFFE) _ "_xFFFE_")
	       (serapeum:string-replace #.(string #\UFFFF) _ "_xFFFF_")))

(defconstant +seconds/day+ #.(* 60 60 24))

(defparameter *1900-base* (local-time:encode-timestamp 0 0 0 0 30 12 1899))

(defparameter *1904-base* (local-time:encode-timestamp 0 0 0 0 1 1 1904))

(defun timestamp-excel (timestamp &optional (base *1900-base*))
  (/ (local-time:timestamp-difference timestamp base) +seconds/day+))

(defun excel-timestamp (days &optional (base *1900-base*))
  (local-time:timestamp+ base (floor (* days +seconds/day+ 1000000000)) :nsec))
