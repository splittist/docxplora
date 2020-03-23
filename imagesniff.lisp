;;;; imagesniff.lisp

(cl:defpackage #:imagesniff
  (:use #:cl)
  (:export
   #:image-type
   #:image-dimensions))

(cl:in-package #:imagesniff)

(defparameter *signatures*
  '((:bmp 0 "image/bmp" "bmp" #(66 77)) ; BM
    (:gif87a 0 "image/gif" "gif" #(71 73 70 56 55 97)) ; GIF87a
    (:gif89a 0 "image/gif" "gif" #(71 73 70 56 57 97)) ; GIF89a
    (:tiff-be 0 "image/tiff" "tif" #(77 77 0 42)) ; MM#\Nul* big-endian
    (:tiff-le 0 "image/tiff" "tif" #(73 73 42 0)) ; II*#\Nul little-endian
    (:png 0 "image/png" "png" #(137 80 78 71 13 10 26 10)) ; #x89,PNG,CRLF,EOF,LF
    (:jfif 6 "image/jpeg" "jpg" #(74 70 73 70)) ; JFIF
    (:exif 6 "image/jpeg" "jpg" #(69 120 105 102)))) ; Exif

(defun get-bmp-info (stream)
  (let (dib-size width height horizontal-resolution vertical-resolution)
    (file-position stream 14)
    (setf dib-size (binary-io:read-value :u32 stream))
    (if (= 12 dib-size) ; ancient OS/2 format
	(setf width (binary-io:read-value :u16 stream)
	      height (binary-io:read-value :u16 stream))
	(setf width (binary-io:read-value :s32 stream)
	      height (binary-io:read-value :s32 stream)))
    (unless (= 12 dib-size)
      (file-position stream 38)
      (setf horizontal-resolution (binary-io:read-value :s32 stream)
	    vertical-resolution (binary-io:read-value :s32 stream)))
    (values width height
	    (if (zerop horizontal-resolution)
		96
		(round (* 0.0254 horizontal-resolution)))
	    (if (zerop vertical-resolution)
		96
		(round (* 0.0254 vertical-resolution))))))

(defun get-gif-info (stream)
  (let (width height)
    (file-position stream 6)
    (setf width (binary-io:read-value :u16 stream)
	  height (binary-io:read-value :u16 stream))
    (values width height 72 72)))

(defun get-tiff-value (stream)
  (let* ((data-type (binary-io:read-value :u16 stream))
	 (data-count (binary-io:read-value :u32 stream)))
    (assert (= 1 data-count))
    (ecase data-type
      (3 (prog1
	     (binary-io:read-value :u16 stream)
	   (binary-io:read-value :u16 stream))) ; skip second word
      (4 (binary-io:read-value :u32 stream))
      (5 (let* ((value-offset (binary-io:read-value :u32 stream))
		(here (file-position stream)))
	   (file-position stream value-offset)
	   (let* ((numerator (binary-io:read-value :u32 stream))
		  (denominator (binary-io:read-value :u32 stream)))
	     (file-position stream here)
	     (/ numerator denominator)))))))

(defun get-tiff-info (stream &key (start 0) (endianness :little))
  (let ((binary-io.common-datatypes:*endianness* endianness))
    (let (ifd0 x-resolution y-resolution image-width image-length (resolution-unit 2))
      (file-position stream (+ 4 start))
      (setf ifd0 (binary-io:read-value :u32 stream))
      (file-position stream ifd0)
      (dotimes (i (binary-io:read-value :u16 stream))
	(file-position stream (+ 2 (* i 12) ifd0))
	(case (binary-io:read-value :u16 stream)
	  (296 (setf resolution-unit (get-tiff-value stream)))
	  (283 (setf y-resolution (get-tiff-value stream)))
	  (282 (setf x-resolution (get-tiff-value stream)))
	  (257 (setf image-length (get-tiff-value stream)))
	  (256 (setf image-width (get-tiff-value stream)))))
      (values image-width image-length
	      (if (null x-resolution)
		  72
		  (ecase resolution-unit
		    (1 72)
		    (2 x-resolution)
		    (3 (round (* 2.54 x-resolution)))))
	      (if (null y-resolution)
		  72
		  (ecase resolution-unit
		    (1 72)
		    (2 y-resolution)
		    (3 (round (* 2.54 y-resolution)))))))))

(defparameter *ihdr* #(73 72 68 82))
(defparameter *phys* #(112 72 89 115))
(defparameter *iend* #(73 69 78 68))

(defun get-png-info (stream)
  (let ((binary-io.common-datatypes:*endianness* :big))
    (let (width height x-resolution y-resolution unit-specifier)
      (file-position stream 8)
      (let* ((ihdr-length (binary-io:read-value :u32 stream))
	     (ihdr-name (binary-io:read-value :vector stream :type :u8 :size 4)))
	(assert (equalp ihdr-name *ihdr*))
	(setf width (binary-io:read-value :u32 stream)
	      height (binary-io:read-value :u32 stream))
	(file-position stream (+ 8 12 ihdr-length)) ; id, length field+name field+crc field
	(loop for prev-position = (file-position stream)
	   for length = (binary-io:read-value :u32 stream)
	   for name = (binary-io:read-value :vector stream :type :u8 :size 4)
	   until (equalp name *iend*)
	   when (equalp name *phys*)
	   do (setf x-resolution (binary-io:read-value :u32 stream)
		    y-resolution (binary-io:read-value :u32 stream)
		    unit-specifier (binary-io:read-value :u8 stream))
	     (loop-finish)
	   do (file-position stream (+ prev-position length 12)))
	(values width height
		(if (null x-resolution)
		    72
		    (if (zerop unit-specifier)
			72
			(round (* x-resolution 0.0254))))
		(if (null y-resolution)
		    72
		    (if (zerop unit-specifier)
			72
			(round (* y-resolution 0.0254)))))))))

(defparameter *eoi* #xFFD9)
(defparameter *app0* #xFFE0)
(defparameter *app1* #xFFE1)

(defun sofp (marker)
  (member marker '(#xFFC0 #xFFC1 #xFFC2 #xFFC3 #xFFC5 #xFFC6 #xFFC7
		   #xFFC9 #xFFCA #xFFCB #xFFCD #xFFCE #xFFCF)))

(defun get-jfif-info (stream)
  (let ((binary-io.common-datatypes:*endianness* :big))
    (let (width height density-units x-density y-density)
      (file-position stream 2) ; skip SOI
      (loop for prev-position = (file-position stream)
	 for marker = (binary-io:read-value :u16 stream)
	 for length = (binary-io:read-value :u16 stream)
	 until (or (equalp marker *eoi*)
		   (and width height density-units x-density y-density))
	 when (and (null width)(sofp marker))
	 do (binary-io:read-value :u8 stream) ; data precision
	   (setf height (binary-io:read-value :u16 stream)
		 width (binary-io:read-value :u16 stream))
	 when (= marker *app0*)
	 do (assert (equalp (binary-io:read-value :vector stream :type :u8 :size 5)
			    #(74 70 73 70 0))) ; JFIF#\Nul
	   (binary-io:read-value :u16 stream) ; version
	   (setf density-units (binary-io:read-value :u8 stream)
		 x-density (binary-io:read-value :u16 stream)
		 y-density (binary-io:read-value :u16 stream))
	 do (file-position stream (+ prev-position 2 length)))
      (values width height
	      (ecase density-units
		(0 72)
		(1 x-density)
		(2 (round (* x-density 2.54))))
	      (ecase density-units
		(0 72)
		(1 y-density)
		(2 (round (* y-density 2.54))))))))

(defun get-exif-info (stream)
  (let ((binary-io.common-datatypes:*endianness* :big))
    (let (width height horizontal-resolution vertical-resolution)
      (file-position stream 2) ; skip SOI
      (loop for prev-position = (file-position stream)
	 for marker = (binary-io:read-value :u16 stream)
	 for length = (binary-io:read-value :u16 stream)
	 until (or (equalp marker *eoi*)
		   (and width height horizontal-resolution vertical-resolution))
	 when (and (null width)(sofp marker))
	 do (binary-io:read-value :u8 stream) ; data precision
	   (setf height (binary-io:read-value :u16 stream)
		 width (binary-io:read-value :u16 stream))
	 when (= marker *app1*)
	 do (assert (equalp (binary-io:read-value :vector stream :type :u8 :size 6)
			    #(69 120 105 102 0 0))) ; Exif#\Nul#\Nul
	   (let* ((tiff-bytes (binary-io:read-value :vector stream
						    :type :u8
						    :size (- length 6))) 
		  (flavour (binary-io:read-value :vector stream :type :u8 :size 4)))
	     (multiple-value-bind (w h hr vr)
		 (babel-streams:with-input-from-sequence
		     (tiff-stream tiff-bytes :element-type '(unsigned-byte 8))
		   (get-tiff-info tiff-stream
				  :endianness (if (= (aref flavour 0) 73) :little :big)))
	       (setf width w height h
		     horizontal-resolution hr
		     vertical-resolution vr)))
	 do (file-position stream (+ prev-position 2 length)))
      (values width height horizontal-resolution vertical-resolution))))

(defgeneric image-type (thing))

(defmethod image-type ((sequence sequence))
  (dolist (candidate *signatures*)
    (destructuring-bind (tag offset mime-type extension bytes) candidate
      (when (alexandria:starts-with-subseq bytes (subseq sequence offset))
	(return (values tag mime-type extension))))))

(defmethod image-type ((stream stream))
  (let ((sequence (make-array 10 :element-type '(unsigned-byte 8) :initial-element 0)))
    (file-position stream 0)
    (read-sequence sequence stream)
    (image-type sequence)))

(defmethod image-type ((pathname pathname))
  (with-open-file (stream pathname :element-type '(unsigned-byte 8))
    (image-type stream)))

(defmethod image-type ((string string))
  (image-type (pathname string)))

(defgeneric image-dimensions (thing))

(defmethod image-dimensions ((stream stream))
  (let ((tag (image-type stream)))
    (case tag
      (:bmp (get-bmp-info stream))
      ((:gif87a :gif89a) (get-gif-info stream))
      (:tiff-be (get-tiff-info stream :endianness :big))
      (:tiff-le (get-tiff-info stream :endianness :little))
      (:png (get-png-info stream))
      (:jfif (get-jfif-info stream))
      (:exif (get-exif-info stream))
      (t nil))))

(defmethod image-dimensions ((pathname pathname))
  (with-open-file (stream pathname :element-type '(unsigned-byte 8))
    (image-dimensions stream)))

(defmethod image-dimensions ((string string))
  (image-dimensions (pathname string)))

(defmethod image-dimensions ((sequence sequence))
  (babel-streams:with-input-from-sequence (stream sequence :element-type '(unsigned-byte 8))
    (image-dimensions stream)))
