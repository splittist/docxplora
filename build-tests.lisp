;;;; build-tests.lisp

(cl:defpackage #:build-tests
  (:use #:cl #:parachute))

(cl:in-package #:build-tests)

(define-test build-suite)
  
(define-test preservep
  :parent build-suite
  (true (docxplora::preservep " "))
  (true (docxplora::preservep " foo"))
  (true (docxplora::preservep "foo "))
  (true (docxplora::preservep " foo "))
  (false (docxplora::preservep "foo"))
  (false (docxplora::preservep "foo bar")))

(defun remove-if/tag-test-helper (tag)
  (let* ((seq-string "<w:r><child1/><child2/><child3 a/><child3 b/></w:r>")
	 (wr (plump:first-child (plump:parse seq-string)))
	 (seq (plump:children wr)))
    (format nil "~{~A~}"
	    (map 'list (alexandria:rcurry #'plump:serialize nil)
		 (docxplora::remove-if/tag tag seq)))))

(define-test  remove-if/tag
  :parent build-suite
  (is string=
      "<child2/><child3 a=\"\"/><child3 b=\"\"/>"
      (remove-if/tag-test-helper "child1"))
  (is string=
      "<child1/><child2/>"
      (remove-if/tag-test-helper "child3")))

(defun remove-if-not/tag-test-helper (tag)
  (let* ((seq-string "<w:r><child1/><child2/><child3 a/><child3 b/></w:r>")
	 (wr (plump:first-child (plump:parse seq-string)))
	 (seq (plump:children wr)))
    (format nil "~{~A~}"
	    (map 'list (alexandria:rcurry #'plump:serialize nil)
		 (docxplora::remove-if-not/tag tag seq)))))

(define-test remove-if-not/tag
  :parent build-suite
  (is string=
      "<child1/>"
      (remove-if-not/tag-test-helper "child1"))
  (is string=
      "<child3 a=\"\"/><child3 b=\"\"/>"
      (remove-if-not/tag-test-helper "child3")))

(defun text-length-test-helper (string)
  (docxplora::text-length
    (plump:first-child (plump:parse string))))

(define-test text-length
  :parent build-suite
  (is =
      0
      (text-length-test-helper "<foo/>"))
  (is =
      1
      (text-length-test-helper "<w:br/>"))
  (is =
      1
      (text-length-test-helper "<w:cr/>"))
  (is =
      1
      (text-length-test-helper "<w:tab/>"))
  (is =
      0
      (text-length-test-helper "<w:t></w:t>"))
  (is =
      1
      (text-length-test-helper "<w:t>1</w:t>"))
  (is =
      5
      (text-length-test-helper "<w:t>12345</w:t>"))
  (is =
      11
      (text-length-test-helper "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"))
  (is =
      10
      (text-length-test-helper
       "<w:p><w:r><w:t>12345</w:t></w:r><w:r><w:t>67890</w:t></w:r></w:p>")))

(defun split-text-element-test-helper (string index)
  (let* ((root (plump:parse string))
	 (wt (plump:first-child root)))
    (docxplora::split-text-element wt index)
    (plump:serialize root nil)))

(define-test split-text-element
  :parent build-suite
  (fail (split-text-element-test-helper "<w:t>1234567890</w:t>" -1))
  (fail (split-text-element-test-helper "<w:t>1234567890</w:t>" 11))
  (is string=
      "<w:t>1234567890</w:t>"
      (split-text-element-test-helper "<w:t>1234567890</w:t>" 0))
  (is string=
      "<w:t>1234567890</w:t>"
      (split-text-element-test-helper "<w:t>1234567890</w:t>" 10))
  (is string=
      "<w:t>12345</w:t><w:t>67890</w:t>"
      (split-text-element-test-helper "<w:t>1234567890</w:t>" 5))
  (is string=
      "<w:t xml:space=\"preserve\">12345 </w:t><w:t>67890</w:t>"
      (split-text-element-test-helper "<w:t>12345 67890</w:t>" 6))
  (is string=
      "<w:t>12345</w:t><w:t xml:space=\"preserve\"> 67890</w:t>"
      (split-text-element-test-helper "<w:t>12345 67890</w:t>" 5))
  (is string=
      "<w:t>12345</w:t><w:t xml:space=\"preserve\">67890 </w:t>"
      (split-text-element-test-helper "<w:t xml:space=\"preserve\">1234567890 </w:t>" 5))
  (is string=
      "<w:t xml:space=\"preserve\"> 12345</w:t><w:t>67890</w:t>"
      (split-text-element-test-helper "<w:t xml:space=\"preserve\"> 1234567890</w:t>" 6)))

(define-test run-from-text-preprocess
  :parent build-suite
  (is equal
      '("foo" :tab "bar")
      (docxplora::run-from-text-preprocess "foo	bar"))
  (is equal
      '(" foo" :tab "bar ")
      (docxplora::run-from-text-preprocess " foo	bar "))
  (is equal
      '("foo" :cr "bar")
      (docxplora::run-from-text-preprocess (format nil "foo~Abar" #\Newline))))

(defun run-from-text-test-helper (string &optional run-properties)
  (plump:serialize (docxplora::run-from-text string run-properties) nil))

(define-test run-from-text
  :parent build-suite
  (is string=
      "<w:r><w:t xml:space=\"preserve\">foo </w:t><w:tab/><w:t>bar</w:t></w:r>"
      (run-from-text-test-helper "foo 	bar"))
  (is string=
      "<w:r><w:rPr><w:style w:val=\"foo\"/></w:rPr><w:t>bar</w:t></w:r>"
      (run-from-text-test-helper
       "bar"
       (wuss:compile-style-to-element '(:r-pr (:style "foo"))))))

(defun child-at-index-test-helper (string index)
  (multiple-value-bind (result sub-index)
      (docxplora::child-at-index
       (plump:first-child
	(plump:parse string))
       index)
    (values
     (and result (plump:serialize result nil))
     sub-index)))

(define-test child-at-index
  :parent build-suite
  (is string=
      "<w:tab/>"
      (child-at-index-test-helper "<w:r><w:tab/></w:r>" 1))
  (is eq
      nil
      (child-at-index-test-helper "<w:r/>" 0))
  (fail
   (child-at-index-test-helper "<w:r/>" -1))
  (fail
   (child-at-index-test-helper "<w:r/>" 1))
  (fail
   (child-at-index-test-helper "<w:r><w:tab/></w:r>" 2))
  (is string=
      "<w:t>12345</w:t>"
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       0))
  (is string=
      "<w:t>12345</w:t>"
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       1))
  (is string=
      "<w:t>12345</w:t>"
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       4))
  (is string=
      "<w:tab/>"
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       5))
  (is string=
      "<w:t>67890</w:t>"
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       6))
  (is string=
      "<w:t>67890</w:t>"
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       11))
  (fail      
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       12))
  (is-values
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       0)
    (string= "<w:t>12345</w:t>")
    (= 0))
  (is-values
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       1)
    (string= "<w:t>12345</w:t>")
    (= 1))
  (is-values
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       2)
    (string= "<w:t>12345</w:t>")
    (= 2))
  (is-values
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       3)
    (string= "<w:t>12345</w:t>")
    (= 3))
  (is-values
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       4)
    (string= "<w:t>12345</w:t>")
    (= 4))
  (is-values
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       5)
    (string= "<w:tab/>")
    (= 0))
  (is-values
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       6)
    (string= "<w:t>67890</w:t>")
    (= 0))
  (is-values
      (child-at-index-test-helper
       "<w:r><w:t>12345</w:t><w:tab/><w:t>67890</w:t></w:r>"
       11)
    (string= "<w:t>67890</w:t>")
    (= 5)))

(defun split-run-test-helper (index string &optional run-properties)
  (let ((run (docxplora::run-from-text string run-properties)))
    (multiple-value-bind (left right)
	(docxplora::split-run run index)
      (values
       (when left (plump:serialize left nil))
       (when right (plump:serialize right nil)))))) 

(define-test split-run
    :parent build-suite
    (is-values
	(split-run-test-helper
	 0
	 "")
      (eq nil)
      (string= "<w:r/>"))
    (is-values
	(split-run-test-helper
	 0
	 "foo")
      (eq nil)
      (string= "<w:r><w:t>foo</w:t></w:r>"))
    (is-values
	(split-run-test-helper
	 3
	 "foo")
      (string= "<w:r><w:t>foo</w:t></w:r>")
      (eq nil))
    (is-values
	(split-run-test-helper
	 3
	 "foo	bar")
      (string= "<w:r><w:t>foo</w:t></w:r>" )
      (string= "<w:r><w:tab/><w:t>bar</w:t></w:r>"))
    (is-values
	(split-run-test-helper
	 4
	 "foo	bar")
      (string= "<w:r><w:t>foo</w:t><w:tab/></w:r>" )
      (string= "<w:r><w:t>bar</w:t></w:r>"))
    (is-values
	(split-run-test-helper
	 1
	 "ab")
      (string= "<w:r><w:t>a</w:t></w:r>")
      (string= "<w:r><w:t>b</w:t></w:r>")))
	
