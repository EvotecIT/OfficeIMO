# JBIG2 decoder fixtures

Source: SerenityOS `Tests/LibGfx/test-inputs/jbig2` at commit
`e70c5343579db8c952c6092e16610e6886dadb24`
(https://github.com/SerenityOS/serenity), BSD-2-Clause, Copyright (c) the SerenityOS developers.
The streams were written by SerenityOS's JBIG2 encoder from JSON descriptions in the same folder; most encode
the 399x400 `json/bitmap.bmp` source image through a different segment type, coding mode or table selection.

`manifest.json` pins each file by SHA-256. `rowsSha256` is the SHA-256 of the decoded page as packed rows with
1 for black and zero padding bits. `reference` names the oracle: `source` is the lossless source image the encoder
was given, and `jbig2dec` is jbig2dec 0.20 (Homebrew) output, used where the expected page is a composite rather
than the source image. jbig2dec was run only as an external reference tool.
