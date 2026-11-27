# Comic archive support

| Source or output | Supported profile |
| --- | --- |
| CBZ | Verified ZIP with JPEG, PNG, GIF or WebP pages, ordered by filename with numeric runs sorted as numbers. |
| CB7 | Unencrypted 7z with Copy, LZMA or LZMA2, including solid folders. Uses the same page order and raster profile as CBZ. See the [7z profile](../OfficeIMO.Archives/SUPPORT.md#7z-reading). |
| ComicInfo | Root `ComicInfo.xml` with title, contributors, publication fields, series, language, Web source, cover and page bookmarks. Page indices refer to the numeric filename order. The first readable page becomes the cover when none is declared. |
| EPUB | EPUB 3 fixed-layout spine with one image per viewport, page bookmarks and cover metadata. |
| PDF | The EPUB renderer uses each viewport as a page. Image decoding and output fidelity follow the PDF renderer's supported profile. |
| HTML and Markdown | Standalone image and bookmark projection. Markdown carries the pictures and publication fields in YAML front matter. |
| CBR | Single-volume, unencrypted RAR 5 with stored pages and ComicInfo. Compressed members raise a diagnostic with the entry name, method and dictionary size. A ZIP or 7z named `.cbr` uses its actual container. See [RAR5 reading](../OfficeIMO.Archives/SUPPORT.md#rar5-reading). |

The reader admits at most 128 MiB of compressed input, 2,000 archive entries,
16 MiB per entry, 120 MiB of expanded data (including solid-folder siblings for CB7), 1,000 page images and
40 million pixels per image. ComicInfo XML is limited to 256 KiB and 16 nesting
levels. Archive paths are limited to 16 segments. Nested archives are outside
the read profile.

Raster formats may have animation, transparency or other features the
destination renderer cannot preserve. The reader imports the page artwork
and does not infer panel boundaries or run OCR.

Info-ZIP, 7-Zip and RAR supplied the independent [comic fixtures](../OfficeIMO.Comic.Tests/Fixtures/README.md). ComicInfo fields follow the [Anansi schema](https://anansi-project.github.io/docs/comicinfo/schemas/v2.0).
