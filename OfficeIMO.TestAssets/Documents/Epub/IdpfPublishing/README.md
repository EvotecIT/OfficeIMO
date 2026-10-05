# Independent EPUB publishing fixture

`childrens-literature` is copied unchanged from
[`IDPF/epub3-samples`, revision `7651e2002b631e6577fadf7e9e0692fa6efb8746`](https://github.com/IDPF/epub3-samples/tree/7651e2002b631e6577fadf7e9e0692fa6efb8746/30/childrens-literature).
The EPUB 3 Samples project is maintained by the W3C EPUB 3 Community Group.
`SOURCE.json` records the revision, file sizes, and SHA-256 hashes. File bytes were
also verified against upstream Git blob hashes when imported.

The sample publication is distributed under
[Creative Commons Attribution-ShareAlike 3.0](https://creativecommons.org/licenses/by-sa/3.0/),
as specified by the upstream [sample table](https://idpf.github.io/epub3-samples/30/samples.html)
and repository license notice. Attribution belongs to the EPUB 3 Samples project
and its contributors; the package retains the named authors, source attribution,
and public-domain notice for the underlying book. No source files were modified.
This fixture license applies to the sample content, independently of OfficeIMO code.

Tests package these files in memory, preserving their contents and writing the
required first, uncompressed `mimetype` entry. The publication exercises two
creator sorting refinements, main/subtitle declarations, nested navigation with
non-link grouping labels, a navigation document in the spine, a hidden page list
covering source pages 169–260, and retained NCX/CSS/image resources. Metadata-edit
tests compare every non-package payload byte before and after editing.

The chapter-split contract first checks that the original scripted publication is
rejected without mutation. It then creates an in-memory static derivative by removing
the navigation script and revealing hidden navigation entries. The independently
produced chapter remains unchanged before splitting; the test checks complete body
text, reading order, print-page targets, and both XHTML and NCX navigation afterward.
The matching merge contract recombines the static derivative and checks complete
body text, the restored spine and all 92 print-page targets.

These are independent-producer read/edit contracts. They do not establish modern
accessibility conformance or independent-reader presentation of the sample.
