# H4 visual acceptance evidence

This folder retains the compact, reviewable output from the advanced held-out HTML rendering gate. The gate ran from clean detached worktrees at source commit `d26d1a3f812b4e116427190b9ba376cff7bdedd6` and passed all eight selected cases on Windows, Linux, and macOS.

The evidence uses report schema 3, acceptance configuration SHA-256 `b347bb8e73ee67bd7130a460de8ee5bd4d3e336ff15cdc9f2756e02a3632a8ed`, and corpus manifest SHA-256 `496f78d459bfd7836987541925f3d6f4b26c87512cd03f319ca99bc4057f67a8`.

## Results

| Platform | Architecture | Runtime | Chromium | PDF rasterizer | Result |
| --- | --- | --- | --- | --- | --- |
| Windows 10.0.26200 | x64 | .NET 10.0.12 | 151.0.7922.34 | Poppler 26.07.0 | 8/8 passed |
| Ubuntu 24.04.3 LTS | x64 | .NET 10.0.12 | 151.0.7922.34 | Poppler 24.02.0 | 8/8 passed |
| macOS 27.0.0 | arm64 | .NET 10.0.12 | 151.0.7922.34 | Poppler 26.09.0 | 8/8 passed |

Each platform folder contains the complete JSON report, the generated Markdown acceptance summary, and the selected OfficeIMO, Chromium, difference, and secondary-reference PNGs used for review:

- [Windows acceptance summary](windows/html-corpus-acceptance.md)
- [Linux acceptance summary](linux/html-corpus-acceptance.md)
- [macOS acceptance summary](macos/html-corpus-acceptance.md)

## Accepted contracts

The gate evaluates three distinct outputs:

- `screen-full-page-v1` compares OfficeIMO screen layout with a Chromium screen screenshot and matching element geometry. Screen markers use exact normalized phrases.
- `print-paged-v1` compares OfficeIMO paged output with Chromium print-to-PDF. Extracted PDF markers use ordered normalized tokens so PDF extraction may interleave layout blocks without accepting arbitrary token order.
- `screen-snapshot-paged-v1` checks fixed-canvas projection against the completed OfficeIMO screen display list. It requires sequential page numbers, uniform 640 x 900 pages, the expected page count, exact final-page padding, declared clipping bounds, full height coverage, and bounded pixel differences.

The capability results bind `static-screen-v1` to both screen intents and `paged-print-v1` to the print intent. They qualify only the selected capabilities and held-out cases recorded in the reports; they do not claim general browser equivalence.

## Known bounded differences

- The professional print catalog intentionally uses centered page comparison because its bleed and printer-mark area differs from the browser sheet. The page background now fills through bleed while the printer-mark reserve remains unpainted.
- The typography specimen admits a bounded cross-block float-layout difference. OfficeIMO currently reserves the float-only formatting context vertically, so the following block starts below it instead of flowing beside it.
- Chromium PDF text extraction varies on Linux for some complex layouts. The source-marker and ordered-token checks remain mandatory on every platform.

The executable runner and report format are documented in the [comparison benchmark guide](../../../../../OfficeIMO.Pdf.Benchmarks.Comparisons/README.md).
