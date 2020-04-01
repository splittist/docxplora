![docxplora](docxploralogo.png)
# docxplora

This is a project to manipulate docx files, primarily those created by Microsoft Word.

* [Open Packing Convention](#opc)
  * [Packages, Parts and Relationships](#packages)
  * [URIs](#uris)
  * [Content Types, Relationship Types and Namespaces](#content-types)
  * [XML](#xml)
* [Open Office XML](#ooxml)
  * [Documents](#documents)
  * [Parts](#parts)
* [WordprocessingML](#wordprocessingml)
  * [Parts](#wml-parts)
  * [Images](#images)
  * [Styles](#styles)
  * [Hyperlinks](#hyperlinks)
  * [Numbering](#nubmering)
  * [Utilities](#utilities)
* [WUSS](#wuss)
  
<a id="opc"></a>
## OPC

[Open Packaging Convention](https://en.wikipedia.org/wiki/Open_Packaging_Conventions/) (OPC) is a container-file format used as the base for modern Microsoft Office documents (as well as some non-Microsoft document formats).

OPC files are zipped containers of other files, and include a map between files and content types (analogous to MIME-types) and 'rels' files to provide a level of indirection for relationships between contained files. The container is a **package**, and the files are **part**s. Each **part** has a **content type**, and may have **relationship**s to other **part**s, each **relationship** having an **Id** and a **relationship type**. **Part**s are named by **uri**s, which are path-like strings.

<a id='packages'></a>
### Packages, Parts and Relationships

*class* **OPC-PACKAGE**

Represents an OPC package.

*function* **MAKE-PACKAGE**

Returns a fresh `OPC-PACKAGE` object.

*function* **OPEN-PACKAGE** `pathname`

Returns an `OPC-PACKAGE` created from the file at `pathname`.

*function* **FLUSH-PACKAGE** `package`

Updates the `/[CONTENT_TYPES].xml` and `*.rels` parts of `package`, and the `CONTENT` of each part in `package`. (See **FLUSH-PART**.)

*function* **SAVE-PACKAGE** `package` &optional `pathname`

Saves `package` to `pathname`, or, if `pathname` is not supplied, the `PACKAGE-PATHNAME` of `package`. Changes `PACKAGE-PATHNAME` to `package`, if supplied. Performs a `FLUSH-PACKAGE` prior to writing to `pathname`.

*function* **PACKAGE-PATHNAME** `package`

The pathname associated with the package, if any.

*class* **OPC-PART**

Represents an opc part within an opc package.

*function* **OPC-PACKAGE** `thing`

The `OPC-PACKAGE` which contains `thing`, an `OPC-PART` or `OPC-RELATIONSHIP`.

*function* **PART-NAME** `part`

The name, a uri, of `part`.

*function* **CONTENT-TYPE** `part`

The content type of `part`.

*function* **CONTENT** `part`

The content of `part` as `(unsigned-byte 8)`s.

*function* **CREATE-PART** `package` `uri` `content-type`

Creates an `OPC-PART` in `package` at `uri` (its name) with content typeb`content-type`. Returns the new `OPC-PART`. `content-type` is a string.

*function* **DELETE-PART** `package` `uri`

Deletes the `OPC-PART` named `uri` from `package`.

*function* **GET-PARTS** `package`

Returns a list of `OPC-PART`s contained in `package`.

*function* **GET-PART** `package` `uri`

Returns the `OPC-PART` named `uri` in `package`.

*class* **OPC-RELATIONSHIP**

Represents a relationship between two parts - the source and the target - having a particular relationship type, and identified by an id. A relationshp is owned by the source (a part or the package; this latter is named "/").

*function* **CREATE-RELATIONSHIP** `source` `uri` `relationship-type` &key `id` `(target-mode "Internal")`

Creates an `OPC-RELATIONSHIP` from `source` (an `OPC-PACKAGE` or `OPC-PART`) to the target at `uri` with relationship type `relationship-type`. If `id` is not supplied, one will be automatically generated. `target-mode` should be "Internal" or "External". `uri`, `relationship-type`, `id` and `target-mode` are strings.

*function* **DELETE-RELATIONSHIP** `source` `id`

Delete the `OPC-RELATIONSHIP` with `id` from `source`, an `OPC-PACKAGE` or `OPC-PART`.

*function* **GET-RELATIONSHIPS** `source`

Returns a list of `OPC-RELATIONSHIP`s owned by `source`, an `OPC-PACKAGE` or `OPC-PART`.

*function* **GET-RELATIONSHIP** `source` `id`

Returns the `OPC-RELATIONSHIP` named `id` owned by `source`, an `OPC-PACKAGE` or `OPC-PART`.

*function* **GET-RELATIONSHIPS-BY-TYPE** `source` `type`

Returns a list of `OPC-RELATIONSHIP`s with the given relationship type `type` (a string), owned by `source`, an `OPC-PACKAGE` or `OPC-PART`.

*function* **SOURCE-URI** `relationship`

The uri (a string) of the owner of the `OPC-RELATIONSHIP` `relationship`.

*function* **TARGET-URI** `relationship`

The uri (a string) of the target (an `OPC-PART` or an external resource) of the `OPC-RELATIONSHIP` `relationship`.

*function* **TARGET-MODE** `relationship`

The target mode (should be "Internal" - indicating an `OPC-PART` in an `OPC-PACKAGE` - or "External" - indicating an external resouce) of the `OPC-RELATIONSHIP` `relationship`.

*function* **RELATIONSHIP-TYPE** `relationship`

The relationship type (a string) of the `OPC-RELATIONSHIP` `relationship`.

*function* **RELATIONSHIP-ID** `relationship`

The unique id (a string) of the relationship (with respect to the source).

<a id='uris'></a>
### URIs

These functions deal only with internal uris, i.e. those indicating `OPC-PART`s within an `OPC-PACKAGE`. Uris are rooted at the package-level with "/", which is sometimes also used to represent the package itself.

*function* **URI-EXTENSION** `uri`

Returns the extension of `uri`, being the string following the last #\..

*function* **URI-FILENAME** `uri`

Returns the filename of `uri`, being the string following the last #\\, or, if no #\\ is found, the whole `uri`.

*function* **URI-DIRECTORY** `uri`

Returns the string up to and including the last #\\ in `uri`, if any.

*function* **URI-RELATIVE** `source` `target`

Returns a uri (a string) identifying `target` relative to `source` (both `target` and `source` being uris). `URI-RELATIVE` and `URI-MERGE` are inverses.

*function* **URI-MERGE** `source` `target`

Takes the uri `target` expressed relative to `source` and returns an absolute uri (all uris being strings). `URI-MERGE` and `URI-RELATIVE` are inverses.

<a id='content-types'></a>
### Content Types, Relationship Types and Namespaces

Because content types, relationship types and namespaces are represented by long strings, convenience functions mapping from shorter strings to the canonical form are provided.

*function* **CT** `name`

Returns the content type represented by the short string `name`.

*function* **RT** `name`

Returns the relationship type represented by the short string `name`.

*function* **NS** `name`

Returns the namespace represented by the short string `name`.

<a id='xml'></a>
### XML

*class* **OPC-XML-PART**

Represents a part in a package with native xml content.

*function* **XML-ROOT** `opc-xml-part`

The **PLUMP:ROOT** of the xml content of the part, if any.

*function* **FLUSH-PART** `part`

Called on each `part` when the enclosing opc package is being flushed, usually prior to saving. An `OPC-XML-PART` will serialize its xml to octets in its `CONTENT`.

*function* **WRITE-PART** `part` `filespec` &key (`if-exists` `:supersede`)

Writes `part` to the file indicated by `filespec`; `if-exists` as for `CL:OPEN-FILE`.

*function* **ENSURE-XML** `part`

Returns `part` as an **OPC-XML-PART**.

*function* **CREATE-XML-PART** `package` `uri` `content-type`

Returns a new **OPC-XML-PART** in the **OPC-PACKAGE** `package` at `uri` (a string) with content type `content-type` (a string).

<a id='ooxml'></a>
## Open Office XML

These classes and functions should, in principle, apply to all Open Office XML documents (Word, Excel and Powerpoint).

<a id='documents'></a>
### Documents

*class* **DOCUMENT**

Represents an **Open Office XML** document.

*function* **OPC-PACKAGE** `document`

Returns the `OPC-PACKAGE` incarnating `document`.

*function* **OPEN-DOCUMENT** `pathname`.

Returns a `DOCUMENT` created from the file at `pathname`.

*function* **MAKE-DOCUMENT**

Returns a fresh `DOCUMENT` object.

*function* **SAVE-DOCUMENT** `document` &optional `pathname`

Saves `document` to `pathname` using `opc:save-package`.

*function* **DOCUMENT-TYPE** `document`

Returns the type of `document` as a keyword: `:wordprocessing-document`, `:spreadsheet-document`, `:presentation-document` or `:opc-package`.

<a id='parts'></a>
### Parts

*function* **GET-PART-BY-NAME** `document` `name` &optional `ensure-xml`

Returns the **OPC-PART** (or subclass) named `name` (a string) in `document`.

*function* **MAIN-DOCUMENT** `document`

Returns the **Main Document** part of the **OOXML** document `document`.

<a id='wordprocessingml'></a>
## WordprocessingML

These classes and functions apply to **WordprocessingML** documents as produced by Microsoft Word and equivalents.

<a id='wml-parts'></a>
### Parts

*function* **COMMENTS** `document`

Returns the **Comments** part of `document`, if any.

*function* **DOCUMENT-SETTINGS** `document`

Returns the **Document Settings* part of `document`, if any.

*function* **ENDNOTES** `document`

Returns the **Endnotes** part of `document`, if any.

*function* **FOOTNOTES** `document`

Returns the **Footnotes** part of `document`, if any.

*function* **FONT-TABLE**

Returns the **Font Table** part of `document`, if any.

*function* **GLOSSARY-DOCUMENT**

Returns the **Glossary Document** part of `document`, if any.

*function* **NUMBERING-DEFINITIONS**

Returns the **Numbering Definitions** part of `document`, if any.

*function* **STYLE-DEFINITIONS** `document`

Returns the **Style Definitions** part of `document`, if any.

*function* **WEB-SETTINGS** `document`

Returns the **Web Settings** part of `document`, if any.

*function* **HEADERS** `document`

Returns a list of **Header** parts of `document`

*function* **FOOTERS**

Returns a list of **Footer** parts of `document`

*function* **ADD-MAIN-DOCUMENT** `document`

Adds a fresh **Main Document** part to `document`, initialized with a `w:body` element containing an empty paragraph (`w:p`), replacing any pre-existing **Main Document** part. Returns the new part.

*function* **ENSURE-MAIN-DOCUMENT** `document`

Returns the **Main Document** part of `document`, either existing or as created by `ADD-MAIN-DOCUMENT`.

<a id='images'></a>
### Images

*function* **MAKE-INLINE-IMAGE** `part` `image`
Adds the image at `image` (a pathname) to the `DOCUMENT` containing `part` (an `OPC-PART`), creates the appropriate **relationship** between `part` and the new image part, and returns a `w:drawing` element suitable for adding to the **xml** in `part`.

<a id='styles'></a>
### Styles

*function* **ADD-STYLE-DEFINITIONS** `document`

Adds a **Style Definitions** part to `document` and returns it.

*function* **ENSURE-STYLE-DEFINITIONS** `document`

Returns the **Style Definitions** part of `document`, either existing or as created by `ADD-STYLE-DEFINITIONS`. 

*function* **ADD-STYLE** `target` `style`

Adds `style` (a `w:style` **PLUMP-DOM:ELMENT**) to `target`, a **document**.

*function* **REMOVE-STYLE** `target` `style`

Removes `style` (a string representing a **style id**, or a `w:style` **PLUMP:ELEMENT**) from `target`, a **document**.

*function* **FIND-STYLE-BY-ID** `target` `style-id` &optional `include-latent`

Returns the **style** from `target` (a **document**) with the **style id** `style-id` (a string), if any. If `include-latent` is non-`NIL`, also looks at **latent styles**.

<a id='hyperlinks'></a>
## Hyperlinks

*function* **ENSURE-HYPERLINK** `source` `uri`

Returns the **OPC:RELATIONSHIP** of type *external* *hyperlink* between `source` (a part or package) and `uri` (a string), creating it if necessary.

<a id='numbering'></a>
### Numbering

*function* **ADD-NUMBERING-DEFINITIONS** `document`

Adds a **Numbering Definitions** part to `document`, and returns it.

*function* **ENSURE-NUMBERING-DEFINITIONS**

Returns the **Numbering Definitions** part of `document`, either existing or as created by **ADD-NUMBERING-DEFINITIONS**.

*function* **MAKE-NUMBERING-START-OVERRIDE** `numbering-definitions` `abstract-num-id` &key (`ilvl` 0) (`start` 1)

Returns a `w:num` **PLUMP-DOM:ELEMENT** with an appropriate **Num Id** for addition to `numbering-definitions` (a **Numbering Definitions** part) referring to `abstract-num-id` (a string or integer) providing for a restart of numbering at **iLvl* `ilvl` at `start` (both integers or strings representing decimal integers).

<a id='utilities'></a>
### Utilities

Some handy (lazy?) shortcuts and generally useful functions.

*function* **FIND-CHILD/TAG** `parent` `child-tag-name`

Returns the (first) child of `parent` (a **PLUMP-DOM:ELEMENT**) with the **tag name** `child-tag-name`.

*function* **FIND-CHILDREN/TAG** `parent` `child-tag-name`

Returns a list of children of `parent` (a **PLUMP-DOM:ELEMENT**) with the **tag name** `child-tag-name`.

*function* **ENSURE-CHILD/TAG** `parent` `child-tag-name` &optional `first`

Returns the child of `parent` (a **PLUMP-DOM:ELMENT**) named `child-tag-name` (a string), creating it if necessary. If `first` is non-`NIL`, a created child will be the first child of `parent`.

*function* **REMOVE-CHILD/TAG** `parent` `child-tag-name`

Removes the child named `child-tag-name` (a string) from `parent` (a **PLUMP-DOM:ELEMENT**), if any.

*function* **MAKE-ELEMENT/ATTRS** `root` `tag-name` &rest `attrs`

Returns a **PLUMP-DOM:ELEMENT** named `tag-name`, a child or `root` (a **PLUMP-DOM:ELEMENT**), with attributes (and values) in the **plist** `attrs`.

*function* **GET-FIRST-ELEMENT-BY-TAG-NAME** `node` `tag`

Returns the first descendant of `node` (a **PLUMP-DOM:ELEMENT**) with the **tag name** `tag` (a string).

*function* **TAGP** `node` `tag`

Does `node` (a **PLUMP-DOM:ELEMENT**) have the *tag name* `tag` (a string)?

*function* **COALESCE-ADJACENT-TEXT** `run`

Transforms `run` (a `w:run` **PLUMP-DOM:ELEMENT**) such that adjacent `w:t` elements are appropriately joined. For example, `<w:r><w:t>foo</w:t><w:t>bar</w:t></w:r>` becomes `<w:r><w:t>foobar</w:t></w:r>`. Useful for simplifying/consolidating programmatically generated text.

*function* **RPR-BOOLEAN-PROPERTY** `rpr` `property-name`

Returns `T` or `NIL` depending on the status of `property-name` (a string) representing a **boolean property** in the given **run properties** `rpr` (a `w:rPr` **PLUMP-DOM:ELEMENT**).

<a id="wuss"></a>
# WUSS

Word-focused Unsatisfactory Style Sheets, a possibly simpler way of specifying the somewhat wordy `w:style` and similar structures.

*function* **COMPILE-STYLE** `style-form`

Returns a string representing the *xml*-form of the *Style* or *Numbering* instance `style-form`, interpreted as follows:

* A **keyword**, representing an element tag

* Zero or more **attribute** *name*-*value* pairs, representing the attributes of the element (but see below)

* Any children, in a parenthesised list.

Symbols are converted from kebab-case to camel-case, strings are left as-is, and other items are `PRINC-TO-STRING`ed. Items in an element or attribute-name position have the "w" namespace prepended. If the length of the attribute list is odd, "w:val" is prepended as a further shorthand. 

Examples might make this less obscure:

```lisp
(:style type "character" custom-style 1 style-id "mdcode"
 (:name "MD Code"
  :ui-priority 1
  :q-format
  :r-pr
  (:no-proof
   :r-fonts ascii "Consolas"
   :sz 20
   :shd 1 color "auto" fill "DCDCDC")))
```

yields:

```xml
<w:style w:type="character" w:customStyle="1" w:styleId="mdcode" >
	<w:name w:val="MD Code" />
	<w:uiPriority w:val="1" />
	<w:qFormat />
	<w:rPr >
		<w:noProof />
		<w:rFonts w:ascii="Consolas" />
		<w:sz w:val="20" />
		<w:shd w:val="1" w:color="auto" w:fill="DCDCDC" />
	</w:rPr>
</w:style>
```
*function* **DECOMPILE-STYLE** `element`

Returns a list of items in the style accepted by **COMPILE-STYLE** generated from `element`, a **PLUMP-DOM:ELEMENT**. The inverse of **COMPILE-STYLE**. The `<w:style>` element generated from the above xml, when fed to **DECOMPILE-STYLE** yields:

```lisp
(:STYLE TYPE "character" CUSTOM-STYLE 1 STYLE-ID "mdcode"
 (:NAME "MD Code" :UI-PRIORITY 1 :Q-FORMAT :R-PR
  (:NO-PROOF :R-FONTS ASCII "Consolas" :SZ 20 :SHD 1 COLOR "auto" FILL
   "DCDCDC")))
```

