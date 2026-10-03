# W3C EPUB spine fixtures

These source publications come from [w3c/epub-tests](https://github.com/w3c/epub-tests)
at commit `54092b4233253e9aac80e93ec4782b380b4b3403`.
Their content files are copied unchanged from `tests/pkg-spine-order-svg` and
`tests/pkg-spine-duplicate-item-rendering`.

The first publication has four SVG spine positions. The second has four reading
positions referencing two unique XHTML resources. Tests package these sources
in memory with a leading uncompressed `mimetype` entry and check the independent
reading-sequence contract.

The fixtures carry W3C copyright and [Software and Document License](https://www.w3.org/copyright/software-license-2015/)
declarations in their package metadata. Preserve those declarations with the files.
