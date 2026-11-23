# Huffman successive-approximation JPEG pairs

Each pair is one 120×72 crop of `OfficeIMO.Examples/Images/Kulek.jpg`, encoded twice
by libjpeg-turbo (as bundled with Magick.NET-Q8 14.16.0) with the same quantization
tables: once baseline and once with the library's standard progressive script.
That script splits every component into spectral bands at reduced precision and then
refines DC and AC down to the last bit, so the final coefficients equal the baseline
ones and both files must decode to the same pixels.

| Pair | Quality | Sampling | Restart interval |
| --- | --- | --- | --- |
| `color-420` | 92 | 2×2, 1×1, 1×1 | none |
| `color-444-restart` | 75 | 1×1, 1×1, 1×1 | 3 MCUs |
| `gray` | 40 | 1×1 | none |

Huffman tables are optimized per file, so refinement scans use their own codes for
zero runs (ZRL) and end-of-band runs (EOBn).
