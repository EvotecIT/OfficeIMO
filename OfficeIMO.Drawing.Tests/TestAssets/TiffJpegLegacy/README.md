# Legacy JPEG-in-TIFF references

These fixtures qualify bounded TIFF compression-6 decoding in the shared raster
owner. `manifest.csv` identifies 114 cases and their source corpora. The generator
recontainers the existing independently encoded JPEG data without recompressing it:
80 baseline table-pointer cases, 16 eight/sixteen-bit lossless table-pointer cases,
six complete lossless interchange images and eight self-contained JPEG strile
cases. Four additional T.81-authored 1×1 lossless scans contain just one entropy
byte, in both TIFF byte orders and sample precisions.

Each `.reference.tif` contains the same compressed image under compression 7.
Managed output must agree exactly between containers. Cases whose encoder changes
Huffman tables between striles retain the first strile of each plane; partial tile
cases retain a smaller visible extent to exercise padding. This recontainering
proves legacy metadata handling with independently encoded entropy, not independent
legacy production. Run `python3 generate.py` from any working directory to rebuild
these cases. `SHA256SUMS` covers TIFFs and native PNG references recursively.

## Upstream files

`upstream/sources.json` records exact LibTIFF v4.7.2 source URLs and SHA-256 values;
`upstream/LICENSE.md` retains its license. The zackthecat file uses a padded tile,
2×2 chroma and integer reference ranges. The chewey file uses 24 strips, 2×1
chroma, a partial interchange header and an SOS header in the first strip.
The unchanged no-rows-per-strip file contains two invalid zero-tag/type/count
records and remains a rejection fixture. Its separately named sanitized copy
removes only those two directory records while preserving payload offsets.

`upstream/generate-native.py` reproduces sanitation and native reference pixels.
Native PNGs were decoded by Pillow 11.3.0 with bundled LibTIFF 4.7.0. Managed RGB
differences are at most 20/255 for zackthecat and the sanitized sample and 14/255
for chewey. Alpha agrees exactly. Chroma interpolation differs between decoders;
these are bounded comparisons, not pixel-identical native rendering claims.
`upstream/normalize-reference.py` reconstructs compression-7 reference files from
the same raw scans and table pointers; managed legacy/reference output agrees
exactly. The source images, sanitized fixture and references are distinct evidence.
Pillow is an opt-in generation/comparison tool and is not a product dependency.

## Boundaries

Complete interchange and self-contained strile JPEGs use the shared strict JPEG
parser. Raw scans require TIFF table pointers and at most four components.
Partial interchange headers without those table tags, mixed chunky lossless
predictors/point transforms, unsupported JPEG precisions/processes and malformed
TIFF directories remain outside the contract. Native Windows acceptance and wider
historical producer coverage are open. CMYK XPS rendering requires an explicit ICC
profile, as it does for the modern TIFF paths.
