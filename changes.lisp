;;;; changes.lisp

(cl:in-package #:docxplora.changes)

(defun fix-character-style (run style-name &optional remove-color)
  (let* ((run-props (ensure-child/tag run "w:rPr" t))
	 (run-style (ensure-child/tag run-props "w:rStyle")))
    (setf (plump:attribute run-style "w:val") style-name)
    (when remove-color
      (remove-child/tag run-props "w:color")))
  run)

(lquery:define-lquery-function fix-character-style (run style-name &optional remove-color)
  (fix-character-style run style-name remove-color))

(defun handle-insertions (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::p" "w::ins" "w::r"
	      (fix-character-style "IPCInsertion")
	      (remove-attr "w:rsidR" "w:rsidRPr"))
    (lquery:$ "w::p" "w::ins"
	      (splice))))

(defun handle-move-tos (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::moveTo" "w::r"
	      (fix-character-style "IPCMoveTo")
	      (remove-attr "w:rsidR" "w:rsidRPr"))
    (lquery:$ "w::moveTo"
	      (splice))
    (lquery:$ "w::moveToRangeStart"
	      (add "w::moveToRangeEnd")
	      (remove))))

(defun handle-move-froms (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::moveFrom" "w::r"
	      (fix-character-style "IPCMoveFrom")
	      (remove-attr "w:rsidR" "w:rsidRPr" "w:rsidDel"))
    (lquery:$ "w::moveFrom"
	      (splice))
    (lquery:$ "w::moveFromRangeStart"
	      (add "w::moveFromRangeEnd")
	      (remove))))

(defun fix-del-text (run)
  (loop for child across (plump:children run)
     when (equal "w:delText" (plump:tag-name child))
     do (setf (plump:tag-name child) "w:t"))
  run)

(lquery:define-lquery-function fix-del-text (run)
  (fix-del-text run))

(defun handle-deletions (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::p" "w::del" "w::r"
	      (fix-character-style "IPCDeletion")
	      (fix-del-text)
	      (remove-attr "w:rsidR" "w:rsidRPr" "w:rsidDel"))
    (lquery:$ "w::p" "w::del"
	      (splice))))

;;; FIXME math run insertions and deletions, and math control stuff

(defun handle-deleted-math-characters (root)
  (lquery:with-master-document (root)
    (lquery:$ "m::r" "w::del"
	      (fix-character-style "IPCDeletion")
	      (splice))))
		
(defun handle-inserted-math-characters (root)
  (lquery:with-master-document (root)
    (lquery:$ "m::r" "w::ins"
	      (fix-character-style "IPCInsertion")
	      (splice))))
  
(defun fix-cell-shading (cell fill)
  (let* ((cell-props (ensure-child/tag cell "w:tcPr"))
	 (cell-shading (ensure-child/tag cell-props "w:shd")))
    (setf (plump:attribute cell-shading "w:val") "1" ;; FIXME ShadingPatternValues.Clear
	  (plump:attribute cell-shading "w:color") "auto"
	  (plump:attribute cell-shading "w:fil") fill))
  cell)

(defun fix-cells (wdel fill)
  (let* ((row (plump:parent (plump:parent wdel)))
	 (cells (find-children/tag row "w:tc")))
    (dolist (cell cells)
      (fix-cell-shading cell fill)))
  wdel)

(lquery:define-lquery-function fix-cells (wdel fill)
  (fix-cells wdel fill))

(defun handle-row-deletions (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::trPr" "w::del"
	      (fix-cells "F7CAAC")
	      (remove))))

(defun handle-row-insertions (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::trPr" "w::ins"
	      (fix-cells "B4C6E7")
	      (remove))))

(defun handle-cell-insertions (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::tcPr" "w::cellIns"
	      (after "<w:shd w:val=\"1\" w:color=\"auto\" w:fill=\"B4C6E7\" />")
	      (remove))))

(defun handle-cell-deletions (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::tcPr" "w::cellDel"
	      (after "<w:shd w:val=\"1\" w:color=\"auto\" w:fill=\"F7CAAC\" />")
	      (remove))))

(defun accept-inserted-and-deleted-paragpraphs (root) ;; TODO insert styled #\Pilcrow_Sign ??
  (lquery:with-master-document (root)
    (lquery:$ "w::pPr" "w::rPr" "w::del" (remove))
    (lquery:$ "w::pPr" "w::rPr" "w::ins" (remove))))

(defun accept-other-changes (root)
  (lquery:with-master-document (root)
    (lquery:$ "w::pPrChange"
	      (add "w::rPrChange")
	      (add "w::tblPrChange")
	      (add "w::tblGridChange")
	      (add "w::trPrChange")
	      (add "w::tcPrChange")
	      (add "w::tblPrChange")
	      (add "w::sectPrChange")
	      (add "w::numberingChange")
	      (add "w::delInstrText") ;; ??
	      (add "w::delText") ;; ??
	      (remove))))

(defun process-part (part)
  (when part
    (let ((root (opc:xml-root part)))
      (handle-inserted-math-characters root)
      (handle-deleted-math-characters root)
      (handle-insertions root)
      (handle-deletions root)
      (handle-move-tos root)
      (handle-move-froms root)
      (handle-row-insertions root)
      (handle-row-deletions root)
      (handle-cell-insertions root)
      (handle-cell-deletions root)
      (accept-inserted-and-deleted-paragpraphs root)
      (accept-other-changes root))))

(defun process-document (document)
  (process-part (main-document document))
  (process-part (endnotes document))
  (process-part (footnotes document))
  (dolist (header (headers document))
    (process-part header))
  (dolist (footer (footers document))
    (process-part footer))
  (accept-styles-changes (style-definitions document)))
    
;;; Styles

(defun accept-styles-changes (style-definitions-part)
  (when style-definitions-part
    (let ((root (opc:xml-root style-definitions-part)))
      (lquery:with-master-document (root)
	(lquery:$ "w::pPrChange"
		  (add "w::rPrChange")
		  (remove))))))

(defun make-ipc-style (id name color fill &key underline strike dstrike)
  (when (and strike dstrike)
    (error "Can't have both single and double strikethrough."))
  (let* ((root (plump:make-root))
	 (style (make-element/attrs root "w:style" "w:type" "character" "w:customStyle" "1" "w:styleId" id)))
    (make-element/attrs style "w:name" "w:val" name)
    (make-element/attrs style "w:basedOn" "w:val" "DefaultParagraphFont")
    (make-element/attrs style "w:uiPriority" "w:val" "1")
    (make-element/attrs style "w:qFormat")
    (let ((run-props (plump:make-element style "w:rPr")))
      (make-element/attrs run-props "w:color" "w:val" color)
      (when fill
	(make-element/attrs run-props "w:shd" "w:val" "1" "w:color" "auto" "w:fill" fill))
      (when underline
	(make-element/attrs run-props "w:u" "w:val" underline))
      (when strike
	(plump:make-element run-props "w:strike"))
      (when dstrike
	(plump:make-element run-props "w:dstrike")))
    style))

(defparameter *classic-ipc-styles*
  (list (make-ipc-style "IPCInsertion" "IPC Insertion" "0000FF" nil :underline "single")
	(make-ipc-style "IPCDeletion" "IPC Deletion" "FF0000" nil :strike t)
	(make-ipc-style "IPCMoveTo" "IPC MoveTo" "00FF00" nil :underline "double")
	(make-ipc-style "IPCMoveFrom" "IPC MoveFrom" "00FF00" nil :dstrike t)))

(defparameter *modern-ipc-styles*
  (list (make-ipc-style "IPCInsertion" "IPC Insertion" "002060" "B4C6E7")
	(make-ipc-style "IPCDeletion" "IPC Deletion" "C00000" "F7CAAC" :strike t)
	(make-ipc-style "IPCMoveTo" "IPC MoveTo" "385623" "C5E0B3")
	(make-ipc-style "IPCMoveFrom" "IPC MoveFrom" "385623" "C5E0B3" :strike t)))

(defun add-ipc-styles (document &optional (style-list *modern-ipc-styles*) replace) ;; FIXME check styles not already there
  (dolist (style style-list)
    (alexandria:if-let (existing-style (find-style-by-id document (plump:attribute style "w:styleId")))
      (when replace
	(remove-style document existing-style)
	(add-style document style))
      (add-style document style))))

;;; pretty changes

(defun pretty-changes (infile &optional outfile)
  (let ((document (open-document infile)))
    (add-ipc-styles document)
    (process-document document)
    (opc:save-package (opc-package document)
		      (or outfile
			  (merge-pathnames (format nil "~A-out" (pathname-name infile))
					   infile)))))
