;;;; wml-namespaces.lisp

(cl:in-package #:docxplora)

(defparameter *md-namespaces*
  '((wpc . "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas")
    (cx . "http://schemas.microsoft.com/office/drawing/2014/chartex")
    (cx1 . "http://schemas.microsoft.com/office/drawing/2015/9/8/chartex")
    (cx2 . "http://schemas.microsoft.com/office/drawing/2015/10/21/chartex")
    (cx3 . "http://schemas.microsoft.com/office/drawing/2016/5/9/chartex")
    (cx4 . "http://schemas.microsoft.com/office/drawing/2016/5/10/chartex")
    (cx5 . "http://schemas.microsoft.com/office/drawing/2016/5/11/chartex")
    (cx6 . "http://schemas.microsoft.com/office/drawing/2016/5/12/chartex")
    (cx7 . "http://schemas.microsoft.com/office/drawing/2016/5/13/chartex")
    (cx8 . "http://schemas.microsoft.com/office/drawing/2016/5/14/chartex")
    (mc . "http://schemas.openxmlformats.org/markup-compatibility/2006")
    (aink . "http://schemas.microsoft.com/office/drawing/2016/ink")
    (am3d . "http://schemas.microsoft.com/office/drawing/2017/model3d")
    (o . "urn:schemas-microsoft-com:office:office")
    (r . "http://schemas.openxmlformats.org/officeDocument/2006/relationships")
    (m . "http://schemas.openxmlformats.org/officeDocument/2006/math")
    (v . "urn:schemas-microsoft-com:vml")
    (wp14 . "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing")
    (wp . "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing")
    (w10 . "urn:schemas-microsoft-com:office:word")
    (w . "http://schemas.openxmlformats.org/wordprocessingml/2006/main")
    (w14 . "http://schemas.microsoft.com/office/word/2010/wordml")
    (w15 . "http://schemas.microsoft.com/office/word/2012/wordml")
    (w16cid . "http://schemas.microsoft.com/office/word/2016/wordml/cid")
    (w16se . "http://schemas.microsoft.com/office/word/2015/wordml/symex")
    (wpg . "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup")
    (wpi . "http://schemas.microsoft.com/office/word/2010/wordprocessingInk")
    (wne . "http://schemas.microsoft.com/office/word/2006/wordml")
    (wps . "http://schemas.microsoft.com/office/word/2010/wordprocessingShape")))

(defparameter *md-ignorable*
  '(("mc:Ignorable" . "w14 w15 w16se w16cid wp14")))