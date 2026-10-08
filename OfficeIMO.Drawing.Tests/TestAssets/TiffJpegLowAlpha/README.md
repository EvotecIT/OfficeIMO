# JPEG TIFF low-alpha fixtures

These 192 35-by-19 fixtures extend the [extra-sample corpus](../TiffJpegAlpha/README.md)
with associated and unassociated alpha. Generate them with that corpus's compiled
LibTIFF/libjpeg-turbo generator and `python3 generate.py <generator> --low-alpha`.
The same producer, five-component scan reconstruction and TIFF metadata limitations
apply. No additional runtime dependencies are introduced.

Source alpha uses 0, 1, 2, 3, 4, 8, 16, 32, 64, 128, 192, 254 and 255 in spatial
bands. JPEG is lossy: the independently decoded alpha samples, rather than those
source values, define the comparison. The matrix covers both byte orders,
strips/tiles, shared/local tables, chunky/separate storage, gray polarities, RGB,
CMYK and centered YCbCr at 1/1 and 2/2 subsampling.

Every decoded alpha sample agrees exactly with the independent reference.
Compositing over black and white differs by at most 3/255 across 127,680 pixels.
Straight RGB can differ by 255/255 near zero alpha because unassociation magnifies
small JPEG/color rounding differences; fully transparent color is not visible.
This is a visible-compositing precision contract, not lossless recovery of source
color or alpha. The `.raw`, `.planes` and `.rgba` files retain the independent
references, and `SHA256SUMS` covers their bytes and the TIFF payloads.

XPS/OpenXPS tests preserve alpha and compare SVG/PDF rendering over black and
white. Native Windows acceptance remains separate from this bounded corpus.
