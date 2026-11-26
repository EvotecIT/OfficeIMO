# FictionBook 2 support

| Source or output | Supported profile |
| --- | --- |
| FB2 | FictionBook 2.0 XML with its declared encoding. DTDs and external entities are refused. |
| FB2.ZIP | One FB2 member in a verified ZIP, including a member in a subdirectory. The reader does not descend into other archives. |
| Contents | Body and nested section titles, section anchors, note links, paragraphs, inline emphasis, poems, quotations, tables and embedded raster images. |
| Metadata | Title, authors, translators, language, identifier, publisher, ISBN, date, subjects and annotation. The referenced cover becomes the EPUB cover image. |
| EPUB and PDF | Reflowable EPUB with linked navigation. PDF uses the EPUB renderer and reports unsupported output features. |
| HTML and Markdown | Standalone output with embedded raster images. Markdown carries the publication fields in YAML front matter. |

The reader admits at most 32 MiB of input and FB2 XML, 128 ZIP entries, 128
body sections at the book level, 1,000 embedded binaries, 250,000 XML reader
nodes and 64 levels of XML nesting. The caller can lower the input byte limit.

Custom stylesheets, series and distribution fields, and unsupported binary types
are reported when encountered. The reader imports FB2 content for conversion
and does not expose an FB2 writer.

The package tests use an independently generated Calibre FB2 and an Info-ZIP
wrapper around it. The [FictionBook schema](https://github.com/gribuser/fb2/blob/master/FictionBook.xsd)
defines the source vocabulary.
