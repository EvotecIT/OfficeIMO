# CCITT uncompressed-mode fixtures

`python3 generate.py` produces 80 specification-authored TIFF fixtures and expected
RGB bytes. These are authored bitstreams, not samples from an independent fax
producer. The generator uses the entry, literal, stuffing and exit codes in
[ITU-T T.4 (1999), Table 5](https://www.itu.int/rec/dologin_pub.asp?id=T-REC-T.4-199904-S!!PDF-E&lang=e&type=items)
and the container rules in [TIFF 6.0, section 11](https://www.itu.int/itudoc/itu-t/com16/tiff-fx/docs/tiff6.pdf).

Each image is 32 by 12 pixels. The matrix covers T.4 one-dimensional and mixed
rows, optional EOL fill, T.6, both byte orders, both fill orders, both grayscale
polarities, strips and tiles. Literal sequences exercise five-zero stuffing,
every zero-to-four-white-pixel exit, both next-run colors, resumed ordinary runs
and a complete literal row whose exit still precedes the next row. Tiles include
bottom padding. Expected pixels come from the generator's source sample arrays.
`SHA256SUMS` covers the TIFF and RGB files.

Core tests compare every pixel and separately exercise vertical/pass resumption,
row-reference use, reserved selectors, row overflow and truncated literal data.
PDF filter tests verify exit consumption before end-of-block markers and both
output polarities. XPS/OpenXPS tests check managed, SVG and PDF rendering.

Independent raw-decoder acceptance remains unqualified. Ghostscript 10.08.0
rejects the extension in its decoder. MuPDF 1.28.2 reports uncompressed fax data
as a format error; the installed OpenJDK 17 TIFF reader fails on the mixed literal
and ordinary-run sample. These failures are not successful interoperability proof.
Independent rendering of OfficeIMO's converted PDF/SVG output is a separate check.
No external decoder is an OfficeIMO runtime dependency.
