# OfficeIMO RTF Golden Corpus

This folder holds small, reviewable fixtures for RTF conversion contracts. Keep fixtures focused and name them after the behavior they protect.

Each fixture should declare:

- source generator or real-world origin when known;
- intended conversion class: `Lossless`, `Semantic`, `Visual`, `Extractive`, or `Diagnostic`;
- target adapters covered: `Rtf`, `Word`, `Html`, `Markdown`, `Reader`, or deferred `Pdf`;
- expected known losses, if any.

PDF fixtures can be referenced in planning, but PDF implementation assertions should wait until the active PDF work stream is ready to integrate.

