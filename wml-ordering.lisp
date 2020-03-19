;;;; wml-ordering.lisp

(cl:in-package #:docxplora)

(defparameter *ppr-order*
          #("w:pStyle"
            "w:keepNext"
            "w:keepLines"
            "w:pageBreakBefore"
            "w:framePr"
            "w:widowControl"
            "w:numPr"
            "w:suppressLineNumbers"
            "w:pBdr"
            "w:shd"
            "w:tabs"
            "w:suppressAutoHyphens"
            "w:kinsoku"
            "w:wordWrap"
            "w:overflowPunct"
            "w:topLinePunct"
            "w:autoSpaceDE"
            "w:autoSpaceDN"
            "w:bidi"
            "w:adjustRightInd"
            "w:snapToGrid"
            "w:spacing"
            "w:ind"
            "w:contextualSpacing"
            "w:mirrorIndents"
            "w:suppressOverlap"
            "w:jc"
            "w:textDirection"
            "w:textAlignment"
            "w:textboxTightWrap"
            "w:outlineLvl"
            "w:divId"
            "w:cnfStyle"
            "w:rPr"
            "w:sectPr"
            "w:pPrChange"))

(defparameter *rpr-order*
         #( "w:moveFrom"
            "w:moveTo"
            "w:ins"
            "w:del"
            "w:rStyle"
            "w:rFonts"
            "w:b"
            "w:bCs"
            "w:i"
            "w:iCs"
            "w:caps"
            "w:smallCaps"
            "w:strike"
            "w:dstrike"
            "w:outline"
            "w:shadow"
            "w:emboss"
            "w:imprint"
            "w:noProof"
            "w:snapToGrid"
            "w:vanish"
            "w:webHidden"
            "w:color"
            "w:spacing"
            "w:_w"
            "w:kern"
            "w:position"
            "w:sz"
            "w14:wShadow"
            "w14:wTextOutline"
            "w14:wTextFill"
            "w14:wScene3d"
            "w14:wProps3d"
            "w:szCs"
            "w:highlight"
            "w:u"
            "w:effect"
            "w:bdr"
            "w:shd"
            "w:fitText"
            "w:vertAlign"
            "w:rtl"
            "w:cs"
            "w:em"
            "w:lang"
            "w:eastAsianLayout"
            "w:specVanish"
            "w:oMath"))

(defparameter *tblpr-order*
          #("w:tblStyle"
            "w:tblpPr"
            "w:tblOverlap"
            "w:bidiVisual"
            "w:tblStyleRowBandSize"
            "w:tblStyleColBandSize"
            "w:tblW"
            "w:jc"
            "w:tblCellSpacing"
            "w:tblInd"
            "w:tblBorders"
            "w:shd"
            "w:tblLayout"
            "w:tblCellMar"
            "w:tblLook"
            "w:tblCaption"
            "w:tblDescription"))

(defparameter *tblborders-order*
          #("w:top"
            "w:left"
            "w:start"
            "w:bottom"
            "w:right"
            "w:end"
            "w:insideH"
            "w:insideV"))

(defparameter *tcpr-order*
          #("w:cnfStyle"
            "w:tcW"
            "w:gridSpan"
            "w:hMerge"
            "w:vMerge"
            "w:tcBorders"
            "w:shd"
            "w:noWrap"
            "w:tcMar"
            "w:textDirection"
            "w:tcFitText"
            "w:vAlign"
            "w:hideMark"
            "w:headers"))

(defparameter *tcborders-order*
          #("w:top"
            "w:start"
            "w:left"
            "w:bottom"
            "w:right"
            "w:end"
            "w:insideH"
            "w:insideV"
            "w:tl2br"
            "w:tr2bl"))

(defparameter *pbdr-order*
          #("w:top"
            "w:left"
            "w:bottom"
            "w:right"
            "w:between"
            "w:bar"))

(defparameter *settings-order*
          #("w:writeProtection" 
            "w:view" 
            "w:zoom" 
            "w:removePersonalInformation" 
            "w:removeDateAndTime" 
            "w:doNotDisplayPageBoundaries" 
            "w:displayBackgroundShape" 
            "w:printPostScriptOverText" 
            "w:printFractionalCharacterWidth" 
            "w:printFormsData" 
            "w:embedTrueTypeFonts" 
            "w:embedSystemFonts" 
            "w:saveSubsetFonts" 
            "w:saveFormsData" 
            "w:mirrorMargins" 
            "w:alignBordersAndEdges" 
            "w:bordersDoNotSurroundHeader" 
            "w:bordersDoNotSurroundFooter" 
            "w:gutterAtTop" 
            "w:hideSpellingErrors" 
            "w:hideGrammaticalErrors" 
            "w:activeWritingStyle" 
            "w:proofState" 
            "w:formsDesign" 
            "w:attachedTemplate" 
            "w:linkStyles" 
            "w:stylePaneFormatFilter" 
            "w:stylePaneSortMethod" 
            "w:documentType" 
            "w:mailMerge" 
            "w:revisionView" 
            "w:trackRevisions" 
            "w:doNotTrackMoves" 
            "w:doNotTrackFormatting" 
            "w:documentProtection" 
            "w:autoFormatOverride" 
            "w:styleLockTheme" 
            "w:styleLockQFSet" 
            "w:defaultTabStop" 
            "w:autoHyphenation" 
            "w:consecutiveHyphenLimit" 
            "w:hyphenationZone" 
            "w:doNotHyphenateCaps" 
            "w:showEnvelope" 
            "w:summaryLength" 
            "w:clickAndTypeStyle" 
            "w:defaultTableStyle" 
            "w:evenAndOddHeaders" 
            "w:bookFoldRevPrinting" 
            "w:bookFoldPrinting" 
            "w:bookFoldPrintingSheets" 
            "w:drawingGridHorizontalSpacing" 
            "w:drawingGridVerticalSpacing" 
            "w:displayHorizontalDrawingGridEvery" 
            "w:displayVerticalDrawingGridEvery" 
            "w:doNotUseMarginsForDrawingGridOrigin" 
            "w:drawingGridHorizontalOrigin" 
            "w:drawingGridVerticalOrigin" 
            "w:doNotShadeFormData" 
            "w:noPunctuationKerning" 
            "w:characterSpacingControl" 
            "w:printTwoOnOne" 
            "w:strictFirstAndLastChars" 
            "w:noLineBreaksAfter" 
            "w:noLineBreaksBefore" 
            "w:savePreviewPicture" 
            "w:doNotValidateAgainstSchema" 
            "w:saveInvalidXml" 
            "w:ignoreMixedContent" 
            "w:alwaysShowPlaceholderText" 
            "w:doNotDemarcateInvalidXml" 
            "w:saveXmlDataOnly" 
            "w:useXSLTWhenSaving" 
            "w:saveThroughXslt" 
            "w:showXMLTags" 
            "w:alwaysMergeEmptyNamespace" 
            "w:updateFields" 
            "w:footnotePr" 
            "w:endnotePr" 
            "w:compat" 
            "w:docVars" 
            "w:rsids" 
            "m.mathPr" 
            "w:attachedSchema" 
            "w:themeFontLang" 
            "w:clrSchemeMapping" 
            "w:doNotIncludeSubdocsInStats" 
            "w:doNotAutoCompressPictures" 
            "w:forceUpgrade" 

            "w:readModeInkLockDown" 
            "w:smartTagType" 

            "w:doNotEmbedSmartTags" 
            "w:decimalSymbol" 
            "w:listSeparator")) 

(defparameter *order-map*
  '(("w:pPr" *ppr-order* t)
    ("w:rPr" *rpr-order* t)
    ("w:tblPr" *tblpr-order* t)
    ("w:tblBorders" *tblborders-order* nil)
    ("w:tcPr" *tcpr-order* t)
    ("w:tcBorders" *tcborders-order* nil)
    ("w:pBdr" *pbdr-order* nil)
    ("w:settings" *settings-order* nil)))

(defun spec-order-sort (element order)
  (setf (plump:children element)
	(sort (plump:children element)
	      #'<
	      :key (lambda (child)
		     (or (position (plump:tag-name child) order :test #'string=)
			 999)))))

(defun spec-order-element (element)
  (let* ((ordering (assoc (plump:tag-name element) *order-map* :test #'string=)))
    (if ordering
	(let ((order (second ordering))
	      (recurse (third ordering)))
	  (spec-order-sort element order)
	  (when recurse
	    (dolist (child (plump:children element))
	      (spec-order-element child))))
	(dolist (child (plump:children element))
	  (spec-order-element child)))))
