# Camera-photo qualification

Add GPS-stripped HEIC/HEIF photographs here, or point `OFFICEIMO_HEIC_REAL_PHOTOS` at a private directory. Keep personal originals outside the repository. Each committed photograph needs its origin, copyright holder, redistribution licence, size and SHA-256 recorded in `../sources.json`.

The intended set is an 8-bit iPhone photo, a 10-bit photo, a rotated photo and a portrait photo with auxiliary depth or matte data. ImageIO platform tests discover `.heic`, `.heif` and `.hif` files and compare the composed SDR base with Apple's full-file decode. An empty slot is reported as a skipped camera-photo qualification, separately from the public-chart tests.

HDR gain-map application, depth reconstruction and thumbnail output are outside the still-image contract.
