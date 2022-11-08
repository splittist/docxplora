(cl:in-package #:docxplora)

(defparameter *wml-content-types*
  '("application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"
    "application/vnd.ms-word.document.macroEnabled.main+xml"
    "application/vnd.ms-word.template.macroEnabledTemplate.main+xml"
    "application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml"))

(defparameter *sml-content-types*
  '("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
    "application/vnd.ms-excel.sheet.macroEnabled.main+xml"
    "application/vnd.ms-excel.template.macroEnabled.main+xml"
    "application/vnd.openxmlformats-officedocument.spreadsheetml.template.main+xml"))

(defparameter *pml-content-types*
  '("application/vnd.openxmlformats-officedocument.presentationml.template.main+xml"
    "application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"
    "application/vnd.ms-powerpoint.template.macroEnabled.main+xml"
    "application/vnd.ms-powerpoint.addin.macroEnabled.main+xml"
    "application/vnd.openxmlformats-officedocument.presentationml.slideshow.main+xml"
    "application/vnd.ms-powerpoint.presentation.macroEnabled.main+xml"))
