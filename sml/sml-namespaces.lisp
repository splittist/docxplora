;;;; sml-namespaces.lisp

(in-package #:sml)

(defparameter *workbook-namespaces*
  '((nil . "http://schemas.openxmlformats.org/spreadsheetml/2006/main")
    ("r" . "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    ("mc" . "http://schemas.openxmlformats.org/markup-compatibility/2006")
    ("x14ac" . "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac")
    ("x15" . "http://schemas.microsoft.com/office/spreadsheetml/2010/11/main")
    ("x16r2" . "http://schemas.microsoft.com/office/spreadsheetml/2015/02/main")
    ("xr" . "http://schemas.microsoft.com/office/spreadsheetml/2014/revision")
    ("xr2" . "http://schemas.microsoft.com/office/spreadsheetml/2015/revision2")
    ("xr3" . "http://schemas.microsoft.com/office/spreadsheetml/2015/revision3")
    ("xr4" . "http://schemas.microsoft.com/office/spreadsheetml/2015/revision4")
    ("xr6" . "http://schemas.microsoft.com/office/spreadsheetml/2016/revision6")
    ("xr10" . "http://schemas.microsoft.com/office/spreadsheetml/2016/revision10")
    ("xlrd" . "http://schemas.microsoft.com/office/spreadsheetml/2017/richdata")
    ("xda" . "http://schemas.microsoft.com/office/spreadsheetml/2017/dynamicarray")
    ))

(defparameter *workbook-ignorables*
  '(("mc:Ignorable" .  "x15 xr xr6 xr10 xr2")))
