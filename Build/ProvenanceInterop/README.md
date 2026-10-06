# C2PA interoperability check

This opt-in test project compares OfficeIMO structural inspection/removal with an independently installed c2patool. It is outside the shipped package graph and the regular test solution. No executable or corpus image is bundled with OfficeIMO.

```sh
python3 Build/ProvenanceInterop/fetch-corpus.py /path/to/task-corpus
dotnet run --project Build/ProvenanceInterop -- /path/to/task-corpus /path/to/trusted/c2patool /path/to/new-evidence.json
```

Use c2patool 0.27.22 for the recorded baseline. Unix hosts also need `setsid` from util-linux (on macOS, `brew install util-linux`). The adapter checks containment before launching the tool. Verification is offline, without caller-supplied trust anchors; the signed case therefore expects `Untrusted`, not `Valid`.

`corpus.json` pins four JPEG files from the [C2PA public test corpus](https://github.com/c2pa-org/public-testfiles/tree/22beccc075707475b038d8789d0136c009e43143/legacy/1.4/image/jpeg), including unmarked, signed, signature-tampered and content-tampered cases. Source SHA-256 checks prevent silently testing replacement fixtures. The images are CC-BY-SA-4.0, with upstream attribution/license copied by the downloader; they are not relicensed as OfficeIMO MIT source. Do not redistribute modified fixtures without meeting their license terms.

The runner writes tool version/hash, platform, corpus revision, source/output hashes, normalized verification findings, pixel equality and case results. It fails on a mismatch and refuses to overwrite existing evidence. Temporary transformed images are deleted after the run. Pixel equality uses OfficeIMO's raster decoder independently of the provenance parser; it is not third-party renderer acceptance. The check does not establish current-spec coverage, certificate-revocation behavior, every format, or cross-platform qualification.
