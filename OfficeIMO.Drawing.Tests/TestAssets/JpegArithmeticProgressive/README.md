# Progressive arithmetic JPEG fixtures

The 150 JPEGs are independently encoded and decoded with libjpeg-turbo 3.2.0.
They cover eight/twelve-bit gray, RGB and YCbCr, 35×19 partial blocks, quality
1/75/100, 1×1/2×1/2×2 chroma, restart intervals of zero/two/three MCUs, high
conditioning-table destinations, and default or L=2/U=5/K=12 conditioning.

Progression 1 uses the library's standard successive approximation script.
Progression 2 uses separate DC scans and two AC spectral bands per component.
Progression 3 initializes those bands at Al=3 and refines each bit down to zero.
Progression 4 emits only DC at Al=3, qualifying a valid coarse preview terminated
by EOI. The other scripts cover interleaved and separate DC, initial AC bands,
DC refinement, signed AC refinement, and statistics reset at scans/restarts.

Native references use integer slow IDCT and nearest/high-quality chroma, with
interblock smoothing disabled. That optional native preview enhancement estimates
missing AC detail; disabling it compares the coefficients present in the stream.
OfficeIMO initializes omitted coefficients to zero. Switching off smoothing changes
only the forty pixel references for the twenty DC-only previews; complete-image
references and all compressed streams remain byte-identical.

Build `generate.c` against libjpeg-turbo headers and link `-ljpeg`; run
`python3 generate.py /absolute/path/to/generate`. The native tool is test-only;
ordinary builds/tests consume retained artifacts and have no native JPEG dependency.
`manifest.csv` records cases and `SHA256SUMS` covers compressed bytes and RGBA8
references. The arithmetic progressive procedures follow
[ITU-T T.81 Annex G](https://www.w3.org/Graphics/JPEG/itu-t81.pdf).

This corpus does not qualify arithmetic CMYK/YCCK, lossless or hierarchical
arithmetic processes, native Windows viewers, or progressive JPEG-in-TIFF. The
TIFF decoder retains its sequential/lossless process contract.
