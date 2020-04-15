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
      (string= "<w:r><w:t>b</w:t></w:r>"))
    (is-values
	(split-run-test-helper
	 4
	 "ab	cd")
      (string= "<w:r><w:t>ab</w:t><w:tab/><w:t>c</w:t></w:r>")
      (string= "<w:r><w:t>d</w:t></w:r>")))

(defun paragraph-insert-text-test-helper (text index &optional run-properties)
  (let ((paragraph
	 (plump:first-child
	  (plump:parse
	   "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>12345</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>67890</w:t></w:r></w:p>"))))
    (plump:serialize
     (docxplora::paragraph-insert-text paragraph text index run-properties)
     nil)))

(defun paragraph-insert-text-test-helper2 (text index &optional run-properties)
  (let ((paragraph
	 (plump:first-child
	  (plump:parse
	   "<w:p><w:r><w:t>1</w:t><w:tab/></w:r><w:r><w:t>abcde</w:t><w:cr/><w:t>567890</w:t></w:r></w:p>"))))
    (plump:serialize
     (docxplora::paragraph-insert-text paragraph text index run-properties)
     nil)))

(define-test paragraph-insert-text
  :parent build-suite
  (is string=
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>12345</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>67890</w:t></w:r><w:r><w:t>foo</w:t></w:r></w:p>"
      (paragraph-insert-text-test-helper "foo" 10))
  (is string=
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>foo</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>12345</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>67890</w:t></w:r></w:p>"
      (paragraph-insert-text-test-helper "foo" 0))
  (is string=
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>12</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>foo</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>345</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>67890</w:t></w:r></w:p>"
      (paragraph-insert-text-test-helper "foo" 2))
  (is string=
      "<w:p><w:r><w:t>1</w:t></w:r><w:r><w:t>FOO</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>abcde</w:t><w:cr/><w:t>567890</w:t></w:r></w:p>"
      (paragraph-insert-text-test-helper2 "FOO" 1))) 

(defun paragraph-delete-text-test-helper (index count)
  (let ((paragraph
	 (plump:first-child
	  (plump:parse
	   "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>12345</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>67890</w:t></w:r></w:p>"))))
    (plump:serialize
     (docxplora::paragraph-delete-text paragraph index count)
     nil)))

(defun paragraph-delete-text-test-helper2 (index count)
  (let ((paragraph
	 (plump:first-child
	  (plump:parse
	   "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>12345</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>ABCDE</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>67890</w:t></w:r></w:p>"))))
    (plump:serialize
     (docxplora::paragraph-delete-text paragraph index count)
     nil)))

(define-test paragraph-delete-text
  :parent build-suite
  (is string=
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>2345</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>67890</w:t></w:r></w:p>"
      (paragraph-delete-text-test-helper 0 1))
  (is string=
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>123</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>890</w:t></w:r></w:p>"
      (paragraph-delete-text-test-helper 3 4))
  (is string=
      "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>123</w:t></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>890</w:t></w:r></w:p>"
      (paragraph-delete-text-test-helper2 3 9))
  (is string=
      "<w:p/>"
      (paragraph-delete-text-test-helper2 0 15)))

(defun paragraph-all-matches-test-helper (regex &key start end
						  match-properties
						  (match-properties-test #'docxplora::element-subsetp))
  (let ((paragraph
	 (plump:first-child
	  (plump:parse
	   "<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>12345</w:t><w:tab/></w:r><w:r><w:t>abcde</w:t><w:cr/></w:r><w:r><w:rPr><w:i/></w:rPr><w:t>567890</w:t></w:r></w:p>"))))
    (docxplora::paragraph-all-matches paragraph regex :start start :end end
				      :match-properties match-properties
				      :match-properties-test match-properties-test)))

(define-test paragraph-all-matches
  :parent build-suite
  (is equal
      '(0 3)
      (paragraph-all-matches-test-helper "123"))
  (is equal
      '(4 5 12 13)
      (paragraph-all-matches-test-helper "5"))
  (is equal
      '(12 13)
      (paragraph-all-matches-test-helper "5" :start 6))
  (is equal
      '(0 5 12 18)
      (paragraph-all-matches-test-helper "\\d+"))
  (is equal
      '(3 7)
      (paragraph-all-matches-test-helper "4.*a"))
  (is equal
      '(11 12)
      (paragraph-all-matches-test-helper "\\n"))
  (is equal
      '(4 5)
      (paragraph-all-matches-test-helper "5" :match-properties '(:r-pr (:b))))
  (is equal
      '(12 13)
      (paragraph-all-matches-test-helper "5" :match-properties (wuss:compile-style-to-element
								'(:r-pr (:i))))))

(defun paragraph-regex-replace-all-test-helper (regex replacement
						&key start end preserve-case simple-calls
						  new-properties match-properties
						  match-properties-test)
  (let ((paragraph
	 (plump:first-child
	  (plump:parse
	   "<w:p><w:r><w:t>12345</w:t><w:tab/></w:r><w:r><w:t>abcde</w:t><w:cr/><w:t>567890</w:t></w:r></w:p>"))))
    (plump:serialize
     (docxplora::paragraph-regex-replace-all paragraph regex replacement
					     :start start :end end
					     :preserve-case preserve-case
					     :simple-calls simple-calls
					     :new-properties new-properties
					     :match-properties match-properties
					     :match-properties-test match-properties-test)
     nil)))

(define-test paragraph-regex-replace-all
  :parent build-suite
  (is string=
      "<w:p><w:r><w:t>XXX</w:t></w:r><w:r><w:t>45</w:t><w:tab/></w:r><w:r><w:t>abcde</w:t><w:cr/><w:t>567890</w:t></w:r></w:p>"
      (paragraph-regex-replace-all-test-helper "123" "XXX"))
  (is string=
      "<w:p><w:r><w:t>FOO</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>abcde</w:t><w:cr/></w:r><w:r><w:t>FOO</w:t></w:r></w:p>"
      (paragraph-regex-replace-all-test-helper "\\d+" "FOO"))
  (is string=
      "<w:p><w:r><w:t>12345</w:t></w:r><w:r><w:t>TAB</w:t></w:r><w:r><w:t>abcde</w:t><w:cr/><w:t>567890</w:t></w:r></w:p>"
      (paragraph-regex-replace-all-test-helper "\\t" "TAB"))
  (is string=
      "<w:p><w:r><w:t>1</w:t></w:r><w:r><w:t>FOO</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>abcde</w:t><w:cr/></w:r><w:r><w:t>FOO</w:t></w:r><w:r><w:t>67890</w:t></w:r></w:p>"
      (paragraph-regex-replace-all-test-helper "[2-5]+" "FOO")))
