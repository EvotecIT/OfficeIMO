# Multi-scan lossless arithmetic TIFF

These 54 TIFFs qualify chunky five-component CMYK plus an extra channel and
chunky four-component 4×2 YCbCr plus an extra channel. They cover eight, twelve
and sixteen bits, associated/straight/unspecified extras, little-endian tiles and
big-endian strips, and centered/cosited YCbCr positioning. Little-endian cases
scan components in forward order; big-endian cases scan them in reverse order.
This is a bounded matrix, not every byte-order/layout combination.

The 243 SOF11 streams are constructed from independently encoded single-component
streams already retained in the planar source corpora. `generate.py` verifies
source hashes, creates a shared frame header, remaps each scan's component ID, and
copies entropy bytes unchanged. Local DAC/DRI definitions remain with their scans;
a zero restart definition before each source's controls prevents state leaking
from an earlier scan. Each scan has one component, so a frame with five components
or with more than ten total sampling units does not exceed the interleaved-scan
limit. Frame geometry is checked against every source component plane.

```sh
python3 generate.py
```

Reference RGBA and the 18 LittleCMS CMYK projections are copied unchanged from the
planar originals. The manifest binds each result to its source file/hash and scan
order. Tests require exact alpha and visible color agreement within 3/255 over
black and white, including explicit-profile CMYK. No runtime dependency or normal
build tool is added.

These are constructed multi-scan files using qualified native sample streams,
not an independent producer corpus for the combined frames. The system LibTIFF
reader rejects all 54 files: 36 extra-channel YCbCr layout errors, six codec
errors, nine component-ID errors and three precision errors. Those diagnostics
do not establish independent full-file acceptance. Native Windows, wider native
TIFF interoperability and default CMYK color interpretation remain unqualified.

The 216 XPS/OpenXPS exports cover 143,640 pixel-center probes per route. MuPDF
1.28.2 renders normalized PDF/SVG within 4/255 and 2/255 respectively, without
warnings. GhostXPS 10.08.0 opens 204 packages but can render blank images (255/255
error), and crashes on 12 twelve-bit CMYK strip exports.
