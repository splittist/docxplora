(cl:in-package #:opc)

(defparameter *relationship-types*
  '(("AUDIO" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" "http://purl.oclc.org/ooxml/officeDocument/relationships/audio")
    ("A_F_CHUNK" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk" "http://purl.oclc.org/ooxml/officeDocument/relationships/aFChunk")
    ("CALC_CHAIN" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain" "http://purl.oclc.org/ooxml/officeDocument/relationships/calcChain")
    ("CERTIFICATE" "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/certificate" "http://purl.oclc.org/ooxml/package/relationships/digital-signature/certificate")
    ("CHART" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" "http://purl.oclc.org/ooxml/officeDocument/relationships/chart")
    ("CHARTSHEET" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet" "http://purl.oclc.org/ooxml/officeDocument/relationships/chartsheet")
    ("CHART_USER_SHAPES" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartUserShapes" "http://purl.oclc.org/ooxml/officeDocument/relationships/chartUserShapes")
    ("COMMENTS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" "http://purl.oclc.org/ooxml/officeDocument/relationships/comments")
    ("COMMENT_AUTHORS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/commentAuthors" "http://purl.oclc.org/ooxml/officeDocument/relationships/commentAuthors")
    ("CONNECTIONS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections" "http://purl.oclc.org/ooxml/officeDocument/relationships/connections")
    ("CONTROL" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" "http://purl.oclc.org/ooxml/officeDocument/relationships/control")
    ("CORE_PROPERTIES" "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" "http://purl.oclc.org/ooxml/package/relationships/metadata/core-properties")
    ("CUSTOM_PROPERTIES" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties" "http://purl.oclc.org/ooxml/officeDocument/relationships/custom-properties")
    ("CUSTOM_PROPERTY" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customProperty" "http://purl.oclc.org/ooxml/officeDocument/relationships/customProperty")
    ("CUSTOM_XML" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" "http://purl.oclc.org/ooxml/officeDocument/relationships/customXml")
    ("CUSTOM_XML_PROPS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps" "http://purl.oclc.org/ooxml/officeDocument/relationships/customXmlProps")
    ("DIAGRAM_COLORS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors" "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramColors")
    ("DIAGRAM_DATA" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData" "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramData")
    ("DIAGRAM_LAYOUT" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout" "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramLayout")
    ("DIAGRAM_QUICK_STYLE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle" "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramQuickStyle")
    ("DIALOGSHEET" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/dialogsheet" "http://purl.oclc.org/ooxml/officeDocument/relationships/dialogsheet")
    ("DRAWING" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" "http://purl.oclc.org/ooxml/officeDocument/relationships/drawing")
    ("ENDNOTES" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes" "http://purl.oclc.org/ooxml/officeDocument/relationships/endnotes")
    ("EXTENDED_PROPERTIES" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" "http://purl.oclc.org/ooxml/officeDocument/relationships/extended-properties")
    ("EXTERNAL_LINK" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink" "http://purl.oclc.org/ooxml/officeDocument/relationships/externalLink")
    ("FONT" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" "http://purl.oclc.org/ooxml/officeDocument/relationships/font")
    ("FONT_TABLE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable" "http://purl.oclc.org/ooxml/officeDocument/relationships/fontTable")
    ("FOOTER" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" "http://purl.oclc.org/ooxml/officeDocument/relationships/footer")
    ("FOOTNOTES" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes" "http://purl.oclc.org/ooxml/officeDocument/relationships/footnotes")
    ("GLOSSARY_DOCUMENT" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument" "http://purl.oclc.org/ooxml/officeDocument/relationships/glossaryDocument")
    ("HANDOUT_MASTER" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/handoutMaster" "http://purl.oclc.org/ooxml/officeDocument/relationships/handoutMaster")
    ("HEADER" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" "http://purl.oclc.org/ooxml/officeDocument/relationships/header")
    ("HYPERLINK" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" "http://purl.oclc.org/ooxml/officeDocument/relationships/hyperlink")
    ("IMAGE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" "http://purl.oclc.org/ooxml/officeDocument/relationships/image")
    ("NOTES_MASTER" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" "http://purl.oclc.org/ooxml/officeDocument/relationships/notesMaster")
    ("NOTES_SLIDE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide" "http://purl.oclc.org/ooxml/officeDocument/relationships/notesSlide")
    ("NUMBERING" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering" "http://purl.oclc.org/ooxml/officeDocument/relationships/numbering")
    ("OFFICE_DOCUMENT" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" "http://purl.oclc.org/ooxml/officeDocument/relationships/officeDocument")
    ("OLE_OBJECT" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/oleObject" "http://purl.oclc.org/ooxml/officeDocument/relationships/oleObject")
    ("ORIGIN" "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin" "http://purl.oclc.org/ooxml/package/relationships/digital-signature/origin")
    ("PACKAGE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" "http://purl.oclc.org/ooxml/officeDocument/relationships/package")
    ("PIVOT_CACHE_DEFINITION" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition" "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotCacheDefinition")
    ("PIVOT_CACHE_RECORDS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/spreadsheetml/pivotCacheRecords" "http://purl.oclc.org/ooxml/officeDocument/relationships/spreadsheetml/pivotCacheRecords")
    ("PIVOT_TABLE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable" "http://purl.oclc.org/ooxml/officeDocument/relationships/pivotTable")
    ("PRES_PROPS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" "http://purl.oclc.org/ooxml/officeDocument/relationships/presProps")
    ("PRINTER_SETTINGS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" "http://purl.oclc.org/ooxml/officeDocument/relationships/printerSettings")
    ("QUERY_TABLE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable" "http://purl.oclc.org/ooxml/officeDocument/relationships/queryTable")
    ("REVISION_HEADERS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionHeaders" "http://purl.oclc.org/ooxml/officeDocument/relationships/revisionHeaders")
    ("REVISION_LOG" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/revisionLog" "http://purl.oclc.org/ooxml/officeDocument/relationships/revisionLog")
    ("SETTINGS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings" "http://purl.oclc.org/ooxml/officeDocument/relationships/settings")
    ("SHARED_STRINGS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" "http://purl.oclc.org/ooxml/officeDocument/relationships/sharedStrings")
    ("SHEET_METADATA" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sheetMetadata" "http://purl.oclc.org/ooxml/officeDocument/relationships/sheetMetadata")
    ("SIGNATURE" "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature" "http://purl.oclc.org/ooxml/package/relationships/digital-signature/signature")
    ("SLIDE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" "http://purl.oclc.org/ooxml/officeDocument/relationships/slide")
    ("SLIDE_LAYOUT" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" "http://purl.oclc.org/ooxml/officeDocument/relationships/slideLayout")
    ("SLIDE_MASTER" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" "http://purl.oclc.org/ooxml/officeDocument/relationships/slideMaster")
    ("SLIDE_UPDATE_INFO" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideUpdateInfo" "http://purl.oclc.org/ooxml/officeDocument/relationships/slideUpdateInfo")
    ("STYLES" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" "http://purl.oclc.org/ooxml/officeDocument/relationships/styles")
    ("TABLE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" "http://purl.oclc.org/ooxml/officeDocument/relationships/table")
    ("TABLE_SINGLE_CELLS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableSingleCells" "http://purl.oclc.org/ooxml/officeDocument/relationships/tableSingleCells")
    ("TABLE_STYLES" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" "http://purl.oclc.org/ooxml/officeDocument/relationships/tableStyles")
    ("TAGS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags" "http://purl.oclc.org/ooxml/officeDocument/relationships/tags")
    ("THEME" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" "http://purl.oclc.org/ooxml/officeDocument/relationships/theme")
    ("THEME_OVERRIDE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/themeOverride" "http://purl.oclc.org/ooxml/officeDocument/relationships/themeOverride")
    ("THUMBNAIL" "http://schemas.openxmlformats.org/package/2006/relationships/metadata/thumbnail" "http://purl.oclc.org/ooxml/package/relationships/metadata/thumbnail")
    ("USERNAMES" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/usernames" "http://purl.oclc.org/ooxml/officeDocument/relationships/usernames")
    ("VIDEO" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" "http://purl.oclc.org/ooxml/officeDocument/relationships/video")
    ("VIEW_PROPS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" "http://purl.oclc.org/ooxml/officeDocument/relationships/viewProps")
    ("VML_DRAWING" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing" "http://purl.oclc.org/ooxml/officeDocument/relationships/vmlDrawing")
    ("VOLATILE_DEPENDENCIES" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/volatileDependencies" "http://purl.oclc.org/ooxml/officeDocument/relationships/volatileDependencies")
    ("WEB_SETTINGS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings" "http://purl.oclc.org/ooxml/officeDocument/relationships/webSettings")
    ("WORKSHEET" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheet")
    ("WORKSHEET_SOURCE" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheetSource" "http://purl.oclc.org/ooxml/officeDocument/relationships/worksheetSource")
    ("XML_MAPS" "http://schemas.openxmlformats.org/officeDocument/2006/relationships/xmlMaps" "http://purl.oclc.org/ooxml/officeDocument/relationships/xmlMaps")))

(defun rt (name &key strict)
  (let ((list (rest (find name *relationship-types* :key #'car :test #'string-equal))))
    (if strict
        (values (second list) (first list))
        (values (first list) (second list)))))

(defun get-relationships-by-type-code (source type-code)
  (multiple-value-bind (transitional strict)
      (rt type-code)
    (append (get-relationships-by-type source transitional)
            (get-relationships-by-type source strict))))
  