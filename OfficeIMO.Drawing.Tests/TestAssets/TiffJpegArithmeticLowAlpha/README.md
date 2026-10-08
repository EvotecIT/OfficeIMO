# Arithmetic JPEG TIFF low alpha

These 192 eight-bit fixtures extend the [arithmetic color/alpha corpus](../TiffJpegArithmeticAlpha/README.md).
They use the same native producer, three-MCU restart interval and TIFF layouts,
with associated/unassociated source alpha bands 0, 1, 2, 3, 4, 8, 16, 32, 64,
128, 192, 254 and 255. Independent decoded alpha is the lossy JPEG reference.

Tests compare visible compositing over black and white within 3/255, with exact
alpha. The 32 CMYK cases additionally use native LittleCMS explicit-profile
references. Straight color close to zero alpha is not claimed as lossless.
Regeneration, digests and full-file/native-consumer limits are documented in the
linked corpus. No additional runtime dependency is required.
