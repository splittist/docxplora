;;;; utils.lisp

(cl:in-package #:docxplora)

#||

360,000 emu / cm; 914,400 emu / in ; 12,700 emu / point

1,440 twip / inch ; 

72 point / inch ;

||#

(defun emu-inch (emu)
  (/ emu 914400))

(defun emu-cm (emu)
  (/ emu 360000))

(defun emu-mm (emu)
  (/ emu 36000))

(defun emu-pt (emu)
  (/ emu 12700))

(defun emu-twip (emu)
  (/ emu 635))

(defun inch-emu (in)
  (* in 914400))

(defun cm-emu (cm)
  (* cm 360000))

(defun mm-emu (mm)
  (* mm 36000))

(defun pt-emu (pt)
  (* pt 12700))

(defun twip-emu (twip)
  (* twip 635))

;; FIXME -- TOO MANY

;;; run properties

(defun boolean-property (pr property-name)
  (let ((prop (first (plump:get-elements-by-tag-name pr property-name))))
    (when prop
      (let ((val (plump:attribute prop "w:val")))
	(cond
	  ((null val) t)
	  ((string= "1" val) t)
	  ((string= "true" val) t)
	  ((string= "on" val) t)
	  ((string= "0" val) nil)
	  ((string= "false" val) nil)
	  ((string= "off" val) nil)
	  (t nil))))))

(defun on-off-attribute (element attribute)
  (multiple-value-bind (attr foundp)
      (plump:attribute element attribute)
    (cond
      ((null foundp) nil)
      ((null attr) t)
      ((string= "1" attr) t)
      ((string= "true" attr) t)
      ((string= "0" attr) nil)
      ((string= "false" attr) nil)
      (t nil)))) ;; FIXME is this correct for the incorrect case?

;;; coalescing elements

(defun coalesce-adjacent-text (run)
  (let ((groups (serapeum:runs (plump:children run) :key #'plump:tag-name :test #'string=)))
    (dolist (group groups)
      (when (and (< 1 (length group))
		 (string= "w:t" (plump:tag-name (alexandria:first-elt group))))
	(let ((first-node (alexandria:first-elt group))
	      (text (serapeum:string-join (map 'list #'plump:text group))))
	  (unless (plump:first-child first-node)
	    (plump:make-text-node first-node)) ; FIXME test if this can create nonsense
	  (setf (plump:text (plump:first-child first-node)) text) ; textual node
	  (when (char= #\Space (alexandria:last-elt text))
	    (setf (plump:attribute first-node "xml:space") "preserve"))
	  (serapeum:do-each (node (subseq group 1))
	    (plump:remove-child node)))))))


#||
(defun node-equal (node1 node2)
  (let ((attrs1 (plump:attributes node1))
	(children1 (plump:children node1))
	(attrs2 (plump:attributes node2))
	(children2 (plump:children node2)))
    (

(defun runs-compatible-p (run1 run2)
  (let ((rpr1 (first (plump:get-elements-by-tag-name run1 "w:rPr")))
	(rpr2 (first (plump:get-elements-by-tag-name run2 "w:rPr"))))
    (and rpr1
	 rpr2
	 (rpr-equal rpr1 rpr2))))
||#


