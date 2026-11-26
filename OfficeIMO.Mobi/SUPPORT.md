# Mobipocket, PalmDOC and Kindle support

| Source or output | Supported profile |
| --- | --- |
| PalmDOC | A Palm `TEXt/REAd` record database. Reads literal Windows-1252 text and preserves line boundaries. Additional bookmark records are reported as omitted. |
| MOBI and AZW | Unencrypted `BOOK/MOBI` versions 1 through 7 with UTF-8 or Windows-1252 text, EXTH metadata, raster resources, embedded index navigation and `filepos` links. |
| AZW3 / KF8 | Unencrypted standalone or dual-format KF8 with skeleton, fragment and navigation indexes, inline raster images and Kindle position links. Separate CSS, SVG and font flows and fixed-layout positioning are reported as lost. |
| Compression | Stored records and PalmDOC compression. HUFF/CDIC-compressed books are refused with a diagnostic. |
| DRM | Encrypted MOBI and Kindle inputs are refused before extracting text, including an encrypted section of a dual-format book. |
| Output | Reflowable EPUB with metadata, embedded raster cover and contents. PDF, standalone HTML and Markdown use the shared publication projection and report destination losses. |

The reader admits at most 64 MiB of input, 16,384 Palm database records,
16 MiB of decompressed text and 4,096 bytes per decompressed text record.
The parser also bounds EXTH items, index records and labels, Kindle
fragments, and HTML nesting. Image loads are restricted to records in the
source publication. External links are kept where supported.

The package covers reading and conversion. MOBI writing, KFX, dictionaries,
Kindle presentation features and DRM decryption are outside the supported
profile. The independently generated fixtures use Calibre. The
[Mobipocket format description](https://wiki.mobileread.com/wiki/MOBI) and
[PalmDOC format description](https://wiki.mobileread.com/wiki/PalmDOC) document
the relevant records.
