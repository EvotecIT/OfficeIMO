# PDF interoperability fixtures

This folder is a deliberately small, checked-in gate rather than a mirror of an
external corpus. Every PDF is pinned by repository commit, upstream path, byte
length, and SHA-256 in `corpus-manifest.json`.

The fixtures come from two public corpora listed by the PDF Association's
[PDF Corpora index](https://github.com/pdf-association/pdf-corpora):

- Open Preservation Foundation `format-corpus`, commit
  `366f068cec399d0cdfd61fa473de3ab6dc858098`, licensed CC0 unless a file says
  otherwise. The selected Cabinet of Horrors files were produced mainly by
  Microsoft Word 2003 and Adobe Acrobat Professional 9.5.2 and cover embedded
  fonts, PDF/A, links, attachments, recoverable corruption, and image failures.
- veraPDF `veraPDF-corpus`, commit
  `49de56cd987929932c9e4fbbbe67d052bf44ef83`, licensed CC BY 4.0. The selected
  atomic PDF/A-1a cases exercise valid ToUnicode character maps.

The suite treats all fixtures as untrusted input, reads them with bounded
OfficeIMO limits, and never executes embedded content. Additions must remain
small, have explicit provenance, and earn a focused contract assertion.
