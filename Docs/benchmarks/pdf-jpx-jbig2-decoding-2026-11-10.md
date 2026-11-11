# JPEG 2000 and JBIG2 decoding in PDF — 2026-11-10

This run validates the managed JPEG 2000 decoder (`OfficeIMO.Core`, `Raster/Jpeg2000`) and the JBIG2 decoder (`OfficeIMO.Pdf`, `Reading/Filters/Jbig2Decoder*`) against independent decoders and encoder source images. Times are single runs of the decoder alone on an Apple-silicon Mac under normal load (Release, .NET 10), after warm-up; allocation is `GC.GetTotalAllocatedBytes` around one decode.

## Method

- JPEG 2000: decoded samples, scaled to eight bits, are compared with `opj_decompress` from OpenJPEG 2.5.4 (Homebrew). The ISO/IEC 15444-4 class-1 reference components (`c1p*.pgx`) from the [openjpeg-data](https://github.com/uclouvain/openjpeg-data/tree/39524bd3a601d90ed8e0177559400d23945f96a9) repository are a second oracle. Those conformance files are download-only because their copyright permits only conformance use.
- JBIG2: decoded pages are compared with `jbig2dec` 0.20 (Homebrew), run on the embedded PDF streams with their `JBIG2Globals`, and with the 399x400 source bitmap that the SerenityOS encoder was given.
- The committed fixtures and their expected hashes are described in [`Images/Jpeg2000/PROVENANCE.md`](../../OfficeIMO.TestAssets/Images/Jpeg2000/PROVENANCE.md) and [`Images/Jbig2/PROVENANCE.md`](../../OfficeIMO.TestAssets/Images/Jbig2/PROVENANCE.md).

## JPEG 2000 results

| Set | Files | Result |
| --- | ---: | --- |
| SerenityOS JasPer, Kakadu and OpenJPEG files | 48 | 39 identical to OpenJPEG; 2 irreversible files within 1. Six differ only because OpenJPEG converts CMYK, Lab and gray-ICC output to RGB. The tile-part-index-overflow file, which OpenJPEG rejects, decodes. |
| ISO/IEC 15444-4 profile 0 and 1 codestreams | 23 | All decode. Lossless streams are identical to OpenJPEG or the class-1 reference components (p0_07 and p0_08 at 12 bits, subsampled p0_02, p0_10 and p1_01, and 257-component p0_13 included). Lossy streams differ by at most 1 except p0_06 component 0 (ROI with irreversible coding): maximum 37 against OpenJPEG and 23 against the conformance reference, where OpenJPEG and the reference themselves differ by 23. |
| JP2 conformance files `file1`–`file9` | 9 | `file1`, `file4`, `file6`, `file9` identical; `file2`/`file3` (sYCC) within 1; `file5`, `file7`, `file8` differ only by OpenJPEG's ICC conversion. |
| JPX images in the PDF corpus | 5 | veraPDF Test Builder, InDesign CS6 (`pdfua-ref-brochure`, 2) identical; LuraDocument scan of NIST SP 800-12 (2) within 1. |

| Image | Size | Time | Allocated |
| --- | --- | ---: | ---: |
| NIST SP 800-12 page 2 background (9/7, ICT) | 2906x3784x3 | 376 ms | 190 MB |
| ISO p0_07 (12-bit, 5/3, RCT) | 2048x2048x3 | 2.15 s | 94 MB |
| InDesign brochure photo (ICC) | 660x901x3 | 497 ms | 13 MB |
| JP2 `file1` | 768x512x3 | 282 ms | 7 MB |
| veraPDF Test Builder image | 640x480x3 | 59 ms | 5 MB |

`opj_decompress` takes 5.0 s, 0.24 s and 1.9 s wall time for p0_07, the brochure photo and the NIST background, including process start and PNM output. Allocation is dominated by one 32-bit coefficient plane per component for the whole tile; single-tile scans therefore hold about 12 bytes per RGB pixel during reconstruction.

## JBIG2 results

| Set | Streams | Result |
| --- | ---: | --- |
| SerenityOS encoder streams (generic, MMR, TPGDON, custom AT, refinement, symbol and text regions with arithmetic and Huffman coding, B.1–B.15 and custom tables, transposed text, halftone and pattern dictionaries, striping, composition operators) | 111 | 110 decode; 80 are identical to jbig2dec and 30 more are identical to the source bitmap where jbig2dec differs (refinement of intermediate regions, refinement in Huffman text regions and aggregate symbols). The unknown-height file was not converted by the harness. |
| JBIG2 images in the PDF corpus (NIST SP 800-12 scan, Danish magazine, NARA JFK scan, veraPDF mixed-raster page) | 304 | All identical to jbig2dec. Median 82 ms, maximum 643 ms per page-size stream (2844x3818), at most 32 MB allocated. |

The largest SerenityOS stream (1.5 MB, Huffman symbol dictionary and text region with thousands of symbols) decodes in 204 ms.

## Coverage before and after

| Area | Before | After |
| --- | --- | --- |
| JPEG 2000 codestream (SIZ, COD/COC, QCD/QCC, RGN, POC, PPM/PPT, tiles and tile-parts, all five progressions, precincts, SOP/EPH, every code-block style, 5/3 and 9/7, RCT/ICT, ROI max-shift, subsampling) | Header validation only; pixels needed a caller-supplied codec | Managed decoder |
| JP2 boxes (`colr` enumerated and ICC, `pclr`, `cmap`, `cdef`) | Opaque Gray/RGB headers only | Applied; boxes misplaced outside `jp2h` are accepted |
| PDF JPXDecode | Pass-through to a caller codec for opaque Gray/RGB; masks, ICC colour spaces and alpha rejected | Opaque Gray/RGB still extract as `.jp2`; everything else normalizes to PNG through the PDF colour, Indexed, mask, `SMaskInData` and output-intent pipeline |
| JBIG2 Huffman symbol dictionaries and text regions, custom tables, collective and uncompressed height-class bitmaps | Rejected | Decoded |
| JBIG2 pattern dictionaries and halftone regions (MMR and arithmetic gray-scale planes, skip, grid vectors) | Rejected | Decoded |
| JBIG2 aggregate symbols with several instances | Decoder lost synchronization | Decoded |

Not covered: JPEG 2000 Part 2 extensions (multiple-component transforms, arbitrary wavelets, JPX compositing), high-throughput (Part 15) codestreams, components above 16 bits, and the JBIG2 generic-region `EXTTEMPLATE` and colour extension segments.
