;;;; images.lisp

(cl:in-package #:docxplora)


(defun next-image-name (part extension)
  (let* ((package (opc:opc-package part))
	 (parts (opc:get-parts package))
	 (partnames (mapcar #'opc:part-name parts))
	 (medianames
	  (remove-if-not (alexandria:curry #'string= "/word/media/")
			 partnames
			 :key #'opc:uri-directory))
	 (imagenames
	  (remove-if-not (alexandria:curry #'string= extension)
			 medianames
			 :key #'opc:uri-extension)))
    (loop for i from 1
       for name = (format nil "/word/media/image~D.~A" i extension)
       while (find name imagenames :test 'string-equal)
	 finally (return name))))

(defun add-image-part (part image) ; FIXME - check duplicates
  (multiple-value-bind (tag mime-type extension)
      (imagesniff:image-type image)
    (unless tag
      (error "Unknown image type for image: ~A" image))
    (let* ((name (next-image-name part extension))
	   (package (opc:opc-package part))
	   (imagepart (opc:create-part package name mime-type)))
      (setf (opc:content imagepart)
	    (alexandria:read-file-into-byte-vector image))
      (let ((rel (opc:create-relationship part
					  (opc:uri-relative (opc:part-name part) name)
					  (opc:rt "IMAGE"))))
	(values imagepart rel)))))

(defun make-inline-image (part image)
  (multiple-value-bind (width height horizontal-resolution vertical-resolution)
      (imagesniff:image-dimensions image)
    (let ((cx (inch-emu (/ width horizontal-resolution)))
	  (cy (inch-emu (/ height vertical-resolution))))
      (multiple-value-bind (imagepart rel)
	  (add-image-part part image)
	(declare (ignore imagepart))
	(make-inline-image-entry (file-namestring image)
				 (opc:relationship-id rel)
				 cx
				 cy)))))

(defun make-inline-image-entry (name rid cx cy)
  (plump:first-child
   (plump:parse
    (format nil
	    "<w:drawing>
  <wp:inline distT=\"0\" distB=\"0\" distL=\"0\" distR=\"0\">
  <wp:extent cx=\"~D\" cy=\"~D\"/>
  <wp:effectExtent l=\"0\" t=\"0\" r=\"0\" b=\"0\"/>
  <wp:docPr id=\"1\" name=\"Picture 1\" descr=\"Description automatically generated\"/>
  <wp:cNvGraphicFramePr>
    <a:graphicFrameLocks
        xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\" noChangeAspect=\"1\"/>
  </wp:cNvGraphicFramePr>
  <a:graphic
      xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">
    <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">
      <pic:pic
          xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">
        <pic:nvPicPr>
          <pic:cNvPr id=\"1\" name=\"~A\"/>
          <pic:cNvPicPr/>
        </pic:nvPicPr>
        <pic:blipFill>
          <a:blip r:embed=\"~A\">
          </a:blip>
          <a:stretch>
            <a:fillRect/>
          </a:stretch>
        </pic:blipFill>
        <pic:spPr>
          <a:xfrm>
            <a:off x=\"0\" y=\"0\"/>
            <a:ext cx=\"~D\" cy=\"~D\"/>
          </a:xfrm>
          <a:prstGeom prst=\"rect\">
            <a:avLst/>
          </a:prstGeom>
        </pic:spPr>
      </pic:pic>
    </a:graphicData>
  </a:graphic>
</wp:inline>
</w:drawing>"
	    cx cy name rid cx cy))))

#+(or)
(defun add-to-md-body (md item)
  (let* ((root (opc:xml-root md))
	 (body (first (plump:get-elements-by-tag-name root "w:body")))
	 (para (plump:make-element body "w:p"))
	 (run (plump:make-element para "w:r")))
    (plump:append-child run item)))

#+(or)
(defun itest ()
  (let* ((doc (make-document))
	 (md (add-main-document doc))
	 (im (make-inline-image md "/home/cabox/workspace/Test/exif-rgb-thumbnail-sony-d700.jpg")))
    (add-to-md-body md im)
    (save-document doc "/home/cabox/workspace/pic.docx")))
