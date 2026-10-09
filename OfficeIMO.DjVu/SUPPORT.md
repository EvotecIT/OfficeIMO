# DjVu support contract

OfficeIMO.DjVu owns managed reading and raster decoding. Reader and PDF packages are thin projections over that owner. DjVu authoring is outside this contract.

| Input or operation | Contract | Limits |
| --- | --- | --- |
| IFF, DJVU, DJVM, DIRM | Single pages, bundled books, and explicit indirect components; stable native page order | Structural bounds, component kinds, aggregate resolved bytes, include depth and cycles are checked. No automatic path or network resolution. |
| TXTa and TXTz | Strict UTF-8 text and nested zones; BZZ-compressed text; byte and UTF-16 offsets | Missing, empty, valid, and corrupt states are distinct. Corrupt text is not implicitly replaced with OCR. |
| NAVM | BZZ outline hierarchy, UTF-8 titles, local component/page targets | External targets remain inert metadata until a PDF URI policy accepts them. Malformed outlines are reported. |
| Sjbz and Djbz | JB2 direct symbols, symbol refinement, placements, shared and inherited dictionaries | Symbols, placements, dictionary chains, bitmap buffers, and arithmetic integer contexts are bounded. Conflicting dictionary branches and the eventual-image-refinement profile are rejected explicitly. |
| Smmr | Whole-image and striped MMR, including inversion | Reuses the managed Core fax codec. Unsupported MMR extensions fail explicitly. |
| BG44, FG44, PM44, BM44 | IW44 version 1.2, grayscale, colour, full and half chroma; progressive BG44/PM44/BM44 and one FG44 chunk | Multiple foreground chunks are invalid under DjVu v3 section 8.3.7 and are rejected. Other IW44 versions fail explicitly. Short image edges have the qualification below. Standalone PM44/BM44 use 100 DPI where INFO is absent. |
| FGbz | Palette colours and BZZ placement correspondence | Palette indices must match decoded placements. |
| BGjp and FGjp | Managed JPEG profiles supported by OfficeIMO.Core | JPEG 2000 BG2k/FG2k fails explicitly, including mixed supported/unsupported layers. No external fallback. |
| Raster output | Native DPI, region selection, gamma, rotation, resolution scaling, caller-owned RGBA | Annotations and viewer settings are reported without active paint. Core resampling at changed DPI is a separate operation from native reconstruction. |
| Reader | Native pages and text geometry, source hashes, optional complete PNG assets and OCR candidates | Text-only by default. Transport requires v11. Stored text has no invented OCR recognition evidence. |
| PDF | Scanned pixels, physical source size, stored searchable text, supported navigation, explicit new OCR | No editable layout/font reconstruction. Missing OCR geometry, corrupt stored text, omitted annotations and outline changes are reported. |

## Resource and ownership boundaries

The default native read profile allows 256 MiB aggregate input, 10,000 pages, 50,000 components, 200,000 chunks, 64 MiB expanded BZZ output, 32 Mi UTF-16 text characters, and one million text zones. A render operation defaults to 32 Mi output pixels and 256 MiB codec/raster working bytes. IW44 decoding permits at most 256 cumulative slices per image and 512 Mi padded coefficient samples across all planes and layers in one render. JB2 permits 128 Mi aggregate decoded symbol samples and 1 MiB decoded comment bytes across the page and inherited dictionaries. Empty symbols and fully clipped placements skip sample traversal. Mask painting permits 128 Mi aggregate clipped bitmap samples, including overlapping placements. Shared text-layer discovery is cached per component; bookmark targets are indexed once. Lower caller limits and cancellation are enforced. Limits describe the owned buffers and counts checked by the APIs; they are not a guarantee about total CLR process memory or execution time.

Inputs, resolved component bytes, and options are copied. Pages and stored text are immutable snapshots; returned raster pixel buffers belong to the caller. Primary input hashes exclude separate indirect component inputs. Applications requiring a complete indirect-document identity must also hash their approved component map.

## Independent qualification

The opt-in [reference runner](../OfficeIMO.DjVu.Verification/README.md) invokes an explicitly supplied DjVuLibre executable outside product packages and normal builds. The published [DjVu specifications](https://djvu.sourceforge.net/doc/) describe the owned codecs; no third-party decoder implementation is vendored or invoked in production.

The [dated raster evidence](../Build/DjVu/Evidence/2026-10-09/rendering-summary.json) records source identities and a complete comparison of 555 pages: *The Time Machine* (232), *The Red Badge of Courage* (252), and the DjVu v3 specification (71). Native-resolution RGB pixels match exactly on every page. Stored text and 86,689 word rectangles match on all 484 archival pages. This is evidence for those identified sources and codec profiles, not universal format coverage.

Checked-in authored fixtures cover progressive/high-frequency colour, grayscale and half chroma, palettes, rotated Unicode text, JPEG, three MMR variants, and independently encoded shared bundled/indirect JB2 dictionaries. Their [provenance manifest](../OfficeIMO.DjVu.Tests/Fixtures/manifest.json) identifies producers, input ownership, and hashes. Normal tests consume fixed fixtures and do not require native tools.

A 13×9 authored colour fixture differs at short IW44 edges by up to five levels in an eight-bit colour channel; its largest per-channel mean absolute difference is 0.471. Short IW44 planes with a dimension below 32 produce `djvu.render.short-iw44-edge` as an `Approximation`, and lossless acceptance rejects it. This measured fixture limit is not a bound promised for every small image. The reference runner retains its stricter maximum-four/mean-0.15 acceptance profile and reports this fixture as outside that profile. JPEG reconstruction can also have decoder rounding differences.

Managed API tests qualify OCR selection, geometry, cancellation, and provenance with a supplied test engine. They do not establish recognition quality for a real OCR model. Complete adapter NativeAOT and browser execution are not qualified; the shared workflow catalog keeps browser availability disabled.
