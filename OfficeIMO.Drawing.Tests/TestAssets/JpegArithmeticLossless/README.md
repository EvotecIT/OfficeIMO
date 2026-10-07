# Lossless arithmetic JPEG fixtures

This corpus exercises SOF11 decoding through the shared managed JPEG owner. It
contains 438 full-resolution grayscale/RGB files and 36 subsampled RGB files,
all 19 × 11 pixels. Source samples use the deterministic expression in
`generate.py`; no third-party photographs are included.

`manifest.csv` covers 420 cases: precisions 2–16, predictors 1–7, point transforms
0 and precision-minus-one, and row-aligned restart intervals. Another 18 files
cover eight/twelve/sixteen-bit grayscale/RGB with no restarts or one/two MCU-row
intervals. `subsampled.csv` covers horizontal, vertical and mixed sampling through
4:1, both point-transform extremes, and restarts. Its `.nearest.rgba` files come
from decoding companion Huffman streams with libjpeg-turbo 3.2.0. This independently
checks the producer's sample prediction and avoids relying on its upsampler.

## Native producer and corrections

The test-only producer is [Thomas Richter's libjpeg](https://github.com/thorfdbg/libjpeg),
commit `c719010a26ce0c666e98b2acf924ad5fc24b4f5d`. It offers a GPLv3 license option;
the executable is built and run separately for fixture generation. Neither its
source implementation nor its executable is linked, shipped, or required by
OfficeIMO or ordinary test runs. The managed decoder follows T.81 Annexes D and H.

The initial 18 files were generated with that revision unchanged. Their bytes are
unchanged when regenerated with the test settings below. Expanded cases use the
small, explicit changes in `prepare_oracle.py`:

- Expose predictor and point-transform selection through the native test driver.
- Allow the lossless scan setup to read the selected predictor tag.
- Correct the initial predictor to `2^(P-Pt-1)`, as required by T.81 H.1.2.1.

The producer originally used `2^(P-1)` even with a point transform. Its own encoder
and decoder agree with each other but disagree with the standard and libjpeg-turbo.
After the correction, 42 Huffman cases covering all seven predictors at 8/12/16 bits
and maximum point transform agree exactly with libjpeg-turbo. Arithmetic cases with
point transforms therefore use a corrected native producer, not an unmodified
independent implementation.

Revision `702114c8130ae818ccd2354cc03e879dac69e579` cannot produce these lossless
files: a later table-selection change requires an absent DQT table. The CLI can
return zero after reporting an error. Generation checks stderr, SOF/SOS parameters,
and decoded sample payloads instead of trusting its exit code alone.

The producer's own decoder also mishandles the bottom edge of the odd-height
subsampled case. The retained subsampled references use libjpeg-turbo's nearest
upsampling of companion Huffman files. They do not qualify a particular smooth
upsampling filter or arithmetic CMYK/YCbCr interpretation.

## Regeneration

Use a separate, disposable checkout of the pinned producer revision. Apply the
settings, then build its executable:

```sh
python3 prepare_oracle.py /path/to/isolated/libjpeg
cd /path/to/isolated/libjpeg
./configure
make -j4 final
```

From this fixture directory, generate with that executable and an existing
libjpeg-turbo `djpeg`. The scratch directory retains intermediate PNM/Huffman
files for inspection and can be removed after verification:

```sh
python3 generate.py /path/to/isolated/libjpeg/jpeg /path/to/djpeg /task/scratch
shasum -a 256 -c SHA256SUMS
```

`generate.py` verifies the native full-resolution decode before accepting each
file. The managed tests check all 99,066 output pixels exactly, preserve native
12/16-bit sample words in both byte orders, and reject invalid scan parameters,
wrong restart ordering, missing termination, cancellation and exhausted memory
budgets. Separate XPS/OpenXPS tests exercise portable SVG/PDF image normalization.

Independent rendering of 948 XPS/OpenXPS exports covers 198,132 pixel-center probes
per route. MuPDF 1.28.2 PDF/SVG output differs by at most 2/255 without warnings.
GhostXPS 10.08.0 opens all exports but differs by up to 255/255, including blank
images. This does not establish Windows XPS Viewer or universal native acceptance.
