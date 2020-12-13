;;;; styles.lisp

(cl:in-package #:docxplora)

(defgeneric add-style (target style)
  (:method ((document document) (style plump-dom:element))
    (assert (tagp  style "w:style") (style) "Need a <w:style>, got ~A")
    (let* ((style-definitions
	    (or (style-definitions document)
		(add-style-definitions document)))
	   (styles (first (plump:get-elements-by-tag-name
			   (opc:xml-root style-definitions)
			   "w:styles"))))
      (plump:append-child styles style))))

(defgeneric remove-style (target style)
  (:method ((document document) (style-id string))
    (alexandria:when-let (style-element (find-style-by-id document style-id))
      (plump:remove-child style-element)))
  (:method ((document document) (style plump:element))
    (assert (tagp style "w:style") (style) "Need a <w:style>, got ~A")
    (plump:remove-child style)))

(defgeneric find-style-by-id (target style-id &optional include-latent) ; FIXME mixes ids and names
  (:method ((document document) (style-id string) &optional include-latent)
    (alexandria:when-let ((style-definitions (style-definitions document)))
      (let ((styles (first (plump:get-elements-by-tag-name
			    (opc:xml-root style-definitions)
			    "w:styles"))))
	(or (find-if (alexandria:curry #'equal style-id)
		     (plump:get-elements-by-tag-name styles "w:style")
		     :key (alexandria:rcurry #'plump:attribute "w:styleId"))
	    (and include-latent
		 (find-if (alexandria:curry #'equal style-id)
			  (plump:get-elements-by-tag-name styles "w:lsdException")
			  :key (alexandria:rcurry #'plump:attribute "w:name"))))))))

(defgeneric style-definitions-styles (target)
  (:method ((style-definitions style-definitions))
    (plump:get-elements-by-tag-name (opc:xml-root style-definitions) "w:style"))
  (:method ((document document))
    (alexandria:when-let ((style-definitions (style-definitions document)))
      (style-definitions-styles style-definitions))))

(defun find-style-definition-by-id (target style-id)
  (let ((styles (style-definitions-styles target)))
    (find-if (alexandria:curry #'string= style-id)
             styles
             :key #'style-id)))

(defun style-type (style)
  (plump:attribute style "w:type"))

(defun style-numbering-style-p (style)
  (equal "numbering" (style-type style)))

(defun style-paragraph-style-p (style)
  (equal "paragraph" (style-type style)))

(defun style-character-style-p (style)
  (equal "character" (style-type style)))

(defun style-table-style-p (style)
  (equal "table" (style-type style)))

(defun style-id (style)
  (plump:attribute style "w:styleId"))

(defun referenced-paragraph-style (paragraph)
  (alexandria:when-let*
      ((ppr (find-child/tag paragraph "w:pPr")))
    (find-child/tag/val ppr "w:pStyle")))

(defun style-numbering-definition-instance-reference (style)
  (serapeum:and-let*
      ((ppr (find-child/tag style "w:pPr"))
       (numpr (find-child/tag ppr "w:numPr")))
    (find-child/tag/val numpr "w:numId")))

#||

Style elements:

__General__

basedOn - val= styleID of parent

link - val=styleID ; links character to paragraph styles and vice versa



name - val=string - primary style name

next - val=styleId ; style for next paragraph; paragraph styles only

qFormat - present or absent - treated as a primary style by application



aliases - val= comma separated list of style names (for UI)

autoRedefine - present or absent (interaction)

hidden - present or absent (UI)

locked - present or absent (UI)

personal - val=boolean (email message context); character styles only

personalCompose - val=boolean (email message context); character styles only

personalReply - ditto

rsid - val=four byte number (Long Hexidecimal Number Value) - editing session

semiHidden - present or absent - initially visible (UI)

uiPriority - val=number optional UI sorting order

unhideWhenUsed - present or absent



__Latent Styles__

latentStyles - behaviour properties (rather than formatting properties) for application (UI etc)
  attributes: count=number, defLockedState=boolean, defQFormat=boolean, defSemiHidden=boolean, defUIPriority=number, defUnhideWhenUsed=boolean, 

lsdException - exception for named latent style
  attributes: name=primary-name (as known to application), locked=boolean, qFormat=boolean, semiHidden=boolean, uiPriority=number, unhideWhenUsed=boolean



__styles.xml__

styles - container for style definitions and latent styles

w:docDefaults>w:rPrDefault>w:rPr^w:pPrDefault>w:pPr

w:latentStyles goes here

style - general style properties; style type (paragraph, character, table, numbering); style type-specific properties
  attributes: customStyle=boolean; default=boolean; styleId=string; type="paragraph" etc

__Character Styles__

type="character"

General Elements

w:rPr

__Paragraph Styles__

type="paragraph"

General Elements

w:rPr - for all runs

w:pPr - paragraph properties; numbering reference is ref to numbering definition only, which in turn has a reference to the paragraph styleon the level associated with the style

__Table Styles__

type="table"

General Elements

w:tblPr

w:tblStylePr w:type="firstRow"|... ; inlcudes w:tblPr {heaps of stuff}

__Numbering Styles__

type="numbering"

w:pPr>w:numPr>w:numId[w:val="~D"

only info specified is ref to numbering defintion

__Run Properties__

b - toggle (bold)

bCs - toggle (complex script bold)

bdo - val="rtl"|"ltr"

bdr - [17.3.2.5]


||#
