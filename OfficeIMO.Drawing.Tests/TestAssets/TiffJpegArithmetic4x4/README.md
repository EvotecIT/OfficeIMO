# Arithmetic JPEG TIFF with 4×4 chroma

These 720 TIFFs cover every precision from 2 through 16 bits, opaque samples,
associated/unassociated full-resolution alpha, centered/cosited 4×4 chroma,
both byte orders, strips and partial tiles, and planar or multi-scan storage.
Images are 5×3 or 17×11. Multi-scan frames use both forward and reverse component
order; interleaving these sampling factors would exceed JPEG's scan MCU limit.

The source planes come from the qualified `TiffJpegChromaPrecision` and
`TiffJpegChromaAlpha` corpora. Libjpeg-turbo 3.2.0 decodes each source plane. The
pinned test-only `thorfdbg/libjpeg` producer independently re-encodes all 1,980
component streams with arithmetic lossless coding, then decodes them again with
exact sample agreement. Predictors 1–7 occur, with point transform zero and
restarts after two component rows. No entropy bytes are altered during assembly.

The TIFFs contain 2,520 streams: planar images retain independent single-component
frames; chunky images combine those scans under a matching SOF11 frame header.
The shared `../tiff_jpeg_multiscan_reference.py` preserves native scan entropy,
conditioning tables and restart intervals. This constructed assembly exercises
multi-scan decoding; it is not an independent whole-frame producer. The native
producer exposes scan lists for progressive JPEG but forces interleaved lossless
scans, so its complete-frame 4×4 color path cannot serve as that oracle.

Device and DCI-P3 profile references retain the source corpora's independently
checked Pillow/LittleCMS values. Opaque ICC references use the corresponding
unassociated-alpha fixture only after verifying byte-identical color-plane JPEG
streams, then replace alpha with 255. All consumed fixtures and references are
checked against their source hashes. The manifest records exact source identity,
predictor, scan order and layout. SHA256SUMS covers every output.

Build the pinned native producer with `../JpegArithmeticLossless/prepare_oracle.py`
and compile `../TiffJpegChroma16/decode.c` against libjpeg-turbo, then run:

```text
python3 generate.py <prepared-jpeg> <compiled-decode> <scratch-directory>
```

All tools are opt-in test dependencies outside the product. The GPLv3 producer is
not vendored or required by OfficeIMO. Independent full-file TIFF/Windows
acceptance, wider profiles and native multi-scan producer evidence remain open.
