# CCITT TIFF decoding fixtures

LibTIFF 4.7.2 independently encodes and decodes these 99 synthetic TIFF images.
They cover Modified Huffman (compression 2), Group 3 one/two-dimensional coding
with and without fill bits (compression 3, options 0/1/4/5), and Group 4
(compression 4). Both byte orders, both photometric polarities, both fill orders,
and strip/tile storage are represented.

Images are 83 by 19 pixels. Strips contain five rows; tiles are 16 by 16.
The pattern includes all-white/all-black rows, long runs requiring makeup codes,
alternating pixels and shifted transitions. Multiple tile rows/columns, partial
edges and nonzero padding test reference-row reset and cropping. `generate.c`
checks every source sample after independent LibTIFF decoding. `manifest.csv`
identifies the encodings and `SHA256SUMS` identifies the exact files.

Three additional single-strip fixtures use the unsigned RowsPerStrip sentinel
`0xFFFFFFFF`. Pass an additional `full` argument to generate these files.

Compile the generator with local LibTIFF headers/library and invoke:
`generate <file> <photometric> <big-endian:0|1> <tiled:0|1> <compression> <options> <fill-order>`.
LibTIFF remains test-only tooling. These cases do not qualify the optional
uncompressed fax extension, native Windows rendering or photographic producers.

Framing and polarity follow [TIFF 6.0 sections 10 and 11](https://www.itu.int/itudoc/itu-t/com16/tiff-fx/docs/tiff6.pdf). Decoding stops at the declared row count; optional uncompressed extensions are rejected.
