# docxplora

### _John Q. Splittist <splittist@splittist.com>_

This is a project to manipulate docx files, primarily those created by
Microsoft Word.

## OPC

[Open Packaging Convention](https://en.wikipedia.org/wiki/Open_Packaging_Conventions/) (OPC) is a container-file format used as the base for modern Microsoft Office documents (as well as some non-Microsoft document formats).

OPC files are zipped containers of other files, and include a map between files and content types (analogous to MIME-types) and 'rels' files to provide a level of indirection for relationships between contained files. The container is a **package**, and the files are **part**s. Each **part** has a **content type**, and may have **relationship**s to other **part**s, each **relationship** having an **Id** and a **relationship type**. **Part**s are named by **uri**s, which are path-like strings.

# Packages, Parts and Relationships

*class* **OPC-PACKAGE**

Represents an OPC package.

*function* **MAKE-PACKAGE**

Returns a fresh `OPC-PACKAGE` object.

*function* **OPEN-PACKAGE** `pathname`

Returns an `OPC-PACKAGE` created from the file at `pathname`.

*function* **FLUSH-PACKAGE** `package`

Updates the `/[CONTENT_TYPES].xml` and `*.rels` parts of `package`, and th e `CONTENT` of each part in `package`. (See **FLUSH-PART**.)

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

*function* **CREATE-RELATIONSHIP** `source` `uri` `relationship-type` &optional `id` `(target-mode "Internal")`

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

# URIs

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

# Content Types, Relationship Types and Namespaces

Because content types, relationship types and namespaces are represented by long strings, convenience functions mapping from shorter strings to the canonical form are provided.

*function* **CT** `name`

Returns the content type represented by the short string `name`.

*function* **RT** `name`

Returns the relationship type represented by the short string `name`.

*function* **NS** `name`

Returns the namespace represented by the short string `name`.

# XML

*class* **OPC-XML-PART**

Represents a part in a package with native xml content.

*function* **XML-ROOT** `opc-xml-part`

The **PLUMP:ROOT** of the xml content of the part, if any.

*function* **FLUSH-PART** `part`

Called on each `part` when the enclosing opc package is being flushed, usually prior to saving. An `OPC-XML-PART` will serialize its xml to octets in its `CONTENT`.

## License

Copyright 2018 John Q. Splittist

Permission is hereby granted, free of charge, to any person obtaining a copy 
of this software and associated documentation files (the "Software"), to deal 
in the Software without restriction, including without limitation the rights 
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell 
copies of the Software, and to permit persons to whom the Software is furnished 
to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included 
in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS 
OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, 
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE 
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER 
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, 
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN 
THE SOFTWARE.

