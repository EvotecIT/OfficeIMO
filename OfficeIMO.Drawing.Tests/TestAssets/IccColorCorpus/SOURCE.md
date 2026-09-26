# ICC raster color corpus

These unmodified profiles cover two RGB matrix/TRC transforms (including wide-gamut DCI-P3), an RGB ICC v4 LUT transform, and a CMYK LUT transform. `reference-srgb.csv` contains packed 8-bit device samples and the sRGB bytes produced by LittleCMS 2.19 through Pillow 12.3.0 with media-relative colorimetric intent and black-point compensation disabled. Run `generate_reference.py` in this directory to reproduce the CSV. The reference is independent of OfficeIMO. Its tolerance in the contract test is two 8-bit channel levels to allow different conforming interpolation and rounding.

| File | Source | SHA-256 |
| --- | --- | --- |
| `icc-dci-p3-matrix.icc` | [ICC DCI-P3-D65 profile](https://registry.color.org/rgb-registry/profiles/DCI-P3-D65.icc) | `36ca884b8eb2ec24675b49e83d84d8eca23500c72f5060276353daf321b2dca3` |
| `littlecms-rgb-matrix.icc` | [LittleCMS `testbed/test5.icc` at `67f272c`](https://github.com/mm2/Little-CMS/blob/67f272c87c31a8073b2f3cdc823ed15ae32e8820/testbed/test5.icc) | `33d771add8e26ad129a5c448ed41da3f8cd08d7e527a3f637ee13220af0afdc7` |
| `littlecms-cmyk-lut.icc` | [LittleCMS `testbed/test1.icc` at `67f272c`](https://github.com/mm2/Little-CMS/blob/67f272c87c31a8073b2f3cdc823ed15ae32e8820/testbed/test1.icc) | `1a996ebe6b5d1e7a21620993ce1903df16e4e8fa6cca312bedfb5ae133e62b4d` |
| `icc-rgb-lut-v4.icc` | [ICC sRGB v4 preference, display class](https://registry.color.org/rgb-registry/profiles/sRGB_v4_ICC_preference_displayclass.icc) | `f54b145a18e4b12112750e672f1c79cac9347dc8403da3955e7f74a352816a21` |

LittleCMS is MIT licensed; its license is included beside the fixtures. The DCI-P3 file's copyright tag identifies the International Color Consortium, whose [profile library terms](https://registry.color.org/profile-library/) permit redistribution. The [sRGB v4 terms](https://registry.color.org/rgb-registry/srgbprofiles) permit redistribution of the unchanged display-class file with its copyright tag retained. The LittleCMS CMYK file identifies itself as a test profile and is used only as a correctness fixture.

Malformed coverage is derived in the test by truncating a valid profile and corrupting a tag offset; configured parser, pixel, and allocation limits are checked separately. The corpus establishes raw device-sample conversion accuracy; it does not establish color-correct extraction from every encoded image format or automatic image optimizer conversion.
