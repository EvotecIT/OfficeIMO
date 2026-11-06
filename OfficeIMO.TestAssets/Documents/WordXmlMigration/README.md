# Native Word XML migration fixtures

Microsoft Word 16.0 build 16.0.20527 produced these two XML serializations from
the same synthetic document on 6 November 2026. The files contain no customer
content. `native-rich-producer.json` records the producer, original package
hash, source XML hashes, image hash, and independently observed document
structure. The paired DOCX hash identifies the producer control; the XML files
are the test inputs retained here.

The document contains custom paragraph and character styles, direct run and
paragraph formatting, a merged table with physical geometry, borders and
shading, one hyperlink, two body image uses, a header table and image, a footer
PAGE field, and a landscape page with explicit margins and two columns. The
same embedded image is referenced by both body pictures and the header.

The WordML 2003 fixture exercises structured migration and explicit loss
reporting. Its legacy settings, list metadata, and other unsupported vocabulary
remain reported losses. The Flat OPC fixture exercises native package
preservation, including auxiliary parts and relationships. These fixtures do
not establish a WordML writer or complete WordML vocabulary coverage.
