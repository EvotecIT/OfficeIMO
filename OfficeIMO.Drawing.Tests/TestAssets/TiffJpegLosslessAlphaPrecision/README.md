# Lossless JPEG-TIFF alpha at additional precisions

These 192 TIFFs qualify Huffman lossless JPEG alpha at 2–7, 9–11 and 13–15 bits.
They complement the existing 8/12/16-bit corpora with both grayscale polarities,
RGB and full-resolution YCbCr. Four paired layouts cover little-endian chunky
and planar strips with associated alpha, and big-endian chunky and planar tiles
with unassociated alpha. The 35×19 image exercises partial strips and edge tiles.
Predictors 1 and 7, separate/interleaved scans, two-row restart intervals and
point transforms zero, one or precision-minus-one are included as recorded in
the manifest. This is a bounded layout matrix, not every option combination.

The shared `../TiffJpegLossless16/generate.c` uses libjpeg-turbo 3.2.0 to encode
and independently decode each compressed segment, verifying every native word
against the source after the point transform. LibTIFF 4.7.2 writes raw strips or
tiles. YCbCr uses an RGB-shaped raw container whose photometric tag is changed;
that declared construction is not independent whole-file YCbCr-alpha producer
acceptance. Full native page words remain in `.raw`; redundant first-segment
sidecars are omitted. The older 12/16-bit generator reproduces all 3,360 existing
fixture/reference files unchanged after its precision extension.

The RGBA reference preserves native color and alpha through unassociation and
fractional YCbCr conversion. The 96 RGB/YCbCr ICC references use LittleCMS 2.19,
the existing DCI-P3 matrix profile and relative colorimetric intent. Alpha is
exact across 127,680 device pixels and 63,840 profiled pixels; color differs by
at most 1/255 and 2/255 respectively. Tests explicitly cover nonzero foreground
color whose native alpha projects to zero at eight bits, protecting conversion
before alpha projection.

Build the shared C generator against test-only libjpeg-turbo and LibTIFF, then
run `python3 generate.py /path/to/generator`. Run `icc-reference.py` with
`LCMS_LIBRARY` pointing to the test-only LittleCMS library if library discovery
is unavailable. `SHA256SUMS` covers all 672 TIFF/native-word/reference files;
regeneration reproduces them exactly. None of these native tools is a product
runtime requirement.

Both XPS dialects retain alpha and visible compositing over black and white
through raster, SVG and PDF readback. The 1,152 device/profile exports have
766,080 pixel-center probes per route. Independent MuPDF 1.28.2 differs by at
most 5/255 in PDF and 2/255 in SVG, without warnings. Four two-bit RGB XPS probes
sample GhostXPS 10.08.0: two device-color cases return zero but differ by 255/255;
two profiled cases terminate with SIGSEGV. The remaining exports were not tested
in GhostXPS. LibTIFF's data-read check rejects 168 containers for precision and
24 chunky YCbCr-alpha containers for layout sizing. Native whole-file acceptance
remains unqualified.

Arithmetic lossless alpha at these additional precisions, subsampled components,
other ICC profiles, and native Windows acceptance remain outside this corpus.
