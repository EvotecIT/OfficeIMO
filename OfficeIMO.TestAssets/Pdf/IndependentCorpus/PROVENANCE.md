# Independent PDF corpus

These PDFs come from producers OfficeIMO does not control. Each file is a byte-identical copy of the upstream file. Its source, SHA-256, license and labels are recorded in [`pdf-corpus-extended.json`](../../../OfficeIMO.Pdf.Benchmarks.Comparisons/Corpus/pdf-corpus-extended.json), and its `sourcePath` points here. `PdfIndependentCorpusTests` checks every file against that manifest and reads it in ordinary test runs. The opt-in corpus runner in `OfficeIMO.Pdf.Benchmarks.Comparisons` uses these copies before downloading.

Files above 2 MB stay download-only. The manifest pins them by URL and SHA-256, but they are not committed:

- `pdfua-ref-magazine-danish` (12.9 MB)
- `pdfua-ref-book-chapter` (2.3 MB)
- `nist-sp800-12-1995-scan` (17.9 MB)
- `rfc9110-xml2rfc` (2.9 MB)
- `naca-report-1135-scan` (6.8 MB)
- `irs-pub17-2023` (3.0 MB)
- `libreoffice-custom-shape-tutorial` (6.3 MB)
- `jp-digital-agency-distiller-slides` (2.3 MB)
- `acl-2023-findings-cjk` (7.4 MB)

The files are test inputs only. No OfficeIMO package includes them, and they keep their own licenses, listed below, rather than the repository's MIT license. Treat them as untrusted input.

## PDF 2.0 examples

Source: [pdf-association/pdf20examples](https://github.com/pdf-association/pdf20examples/tree/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa), commit `c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa`.

License: [CC BY-SA 4.0](https://creativecommons.org/licenses/by-sa/4.0/), as stated in the repository's [LICENSE.md](https://github.com/pdf-association/pdf20examples/blob/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa/LICENSE.md).

Attribution: PDF 2.0 example files © PDF Association, from https://github.com/pdf-association/pdf20examples, licensed under CC BY-SA 4.0. The files are unmodified.

| File | Source | SHA-256 | Producer |
| --- | --- | --- | --- |
| `pdf20-simple.pdf` | [source](https://raw.githubusercontent.com/pdf-association/pdf20examples/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa/Simple%20PDF%202.0%20file.pdf) | `296d2a0b2ce19b606f29265694f194a754fbc61783982b5b8d730e8637482236` | PDF Association PDF 2.0 examples (hand-written) |
| `pdf20-utf8-annotation.pdf` | [source](https://raw.githubusercontent.com/pdf-association/pdf20examples/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa/PDF%202.0%20UTF-8%20string%20and%20annotation.pdf) | `f3a018213564a420db81062a0e584fbcc5f705e353ed477da43bab6222f137b6` | PDF Association PDF 2.0 examples (hand-written) |
| `pdf20-image-bpc.pdf` | [source](https://raw.githubusercontent.com/pdf-association/pdf20examples/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa/PDF%202.0%20image%20with%20BPC.pdf) | `fc517dd0d4c72d1b11e3449acc3e816d2719dc14e69a7e14b84da02808103370` | PDF Association PDF 2.0 examples (hand-written) |
| `pdf20-incremental-save.pdf` | [source](https://raw.githubusercontent.com/pdf-association/pdf20examples/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa/PDF%202.0%20via%20incremental%20save.pdf) | `9b25d59d07034634a990e79db6ef82b9f86769d9b94d8d8635f20d06aaec9d18` | PDF Association PDF 2.0 examples (hand-written) |
| `pdf20-offset-start.pdf` | [source](https://raw.githubusercontent.com/pdf-association/pdf20examples/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa/PDF%202.0%20with%20offset%20start.pdf) | `a5b92313770f16ba1db649ccdc97d42f7744659c691897538734eebdc1bc9e9a` | PDF Association PDF 2.0 examples (hand-written) |
| `pdf20-page-output-intent.pdf` | [source](https://raw.githubusercontent.com/pdf-association/pdf20examples/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa/PDF%202.0%20with%20page%20level%20output%20intent.pdf) | `7c13afe3078a4c202184edce3169f7bc842c8fd3db9da7140af6bc0b3b2fa44a` | PDF Association PDF 2.0 examples (hand-written) |
| `pdf20-utf8-test.pdf` | [source](https://raw.githubusercontent.com/pdf-association/pdf20examples/c20f2c17bfcc4baab7cfe62e70fae64caf14d5fa/pdf20-utf8-test.pdf) | `a7ca717ef062df0a65e58487d26358c81424b76084a2000ca1c74b9e5776b7a1` | PDF Association PDF 2.0 examples (hand-written) |

## PDF/UA Reference Suite

Source: the [PDF/UA Reference Suite](https://pdfa.org/resource/pdfua-reference-suite/) by the PDF Association, copied from [veraPDF/veraPDF-regression-tests](https://github.com/veraPDF/veraPDF-regression-tests/tree/0fabb1c44f53911d0680c9d8743c5ec3c95dc52d/WCAG_2_2/1.7) commit `0fabb1c44f53911d0680c9d8743c5ec3c95dc52d`.

License: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/).

Attribution: each work below is part of the PDF/UA Reference Suite, © PDF Association and the named authors, licensed under CC BY 4.0. The files are unmodified.

| File | Work | Source | SHA-256 | Producer |
| --- | --- | --- | --- | --- |
| `pdfua-ref-academic-abstract.pdf` | "More than just digital paper—hidden features of the PDF format", Dietrich von Seggern, Klaas Posselt, Tamir Hassan and Thomas Zellmann | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-regression-tests/0fabb1c44f53911d0680c9d8743c5ec3c95dc52d/WCAG_2_2/1.7/PDFUA-Ref-2-03_AcademicAbstract.pdf) | `e61257d371261867926eb006f8895569a95fcfbcb93fca3be04b8b9187cb93a0` | pdfTeX 1.40.18 (LaTeX acmart) |
| `pdfua-ref-presentation.pdf` | "DocEng 2019: PDF-Workshop, Tagged PDF slides", PDF Association | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-regression-tests/0fabb1c44f53911d0680c9d8743c5ec3c95dc52d/WCAG_2_2/1.7/PDFUA-Ref-2-04_Presentation.pdf) | `90acb829ca8bfa037796592539d15010e9099d1cd2622aac1fce7d9fd3ae7575` | Adobe InDesign 15.0 |
| `pdfua-ref-book-chapter-german.pdf` | "Wie kommen Sie zu einem barrierefreien PDF? – Grundüberlegungen", Klaas Posselt | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-regression-tests/0fabb1c44f53911d0680c9d8743c5ec3c95dc52d/WCAG_2_2/1.7/PDFUA-Ref-2-05_BookChapter-german.pdf) | `fc525b8899da31779b6143c16053779cd560d4417018a901472de33c71252352` | Adobe InDesign CC 13.1 |
| `pdfua-ref-brochure.pdf` | "PDF/UA – ISO 14289: The new standard for universally accessible PDF documents and forms", PDF/UA Competence Center | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-regression-tests/0fabb1c44f53911d0680c9d8743c5ec3c95dc52d/WCAG_2_2/1.7/PDFUA-Ref-2-06_Brochure.pdf) | `aefd5ab48a908b9996ffadcb14d8ff31a2c653657cbe011a46f4cf61396a589e` | Adobe InDesign CS6 |
| `pdfua-ref-formal-document.pdf` | "Concepts of Abuse and Neglect", Rutgers University | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-regression-tests/0fabb1c44f53911d0680c9d8743c5ec3c95dc52d/WCAG_2_2/1.7/PDFUA-Ref-2-07_FormalDocument.pdf) | `14aecf649931454321ff3b13c6014422bb06c955d7274281a2f7a922458c281f` | AbleDocs remediation |
| `pdfua-ref-form.pdf` | "PDF/UA form example", PDF Association / Klaas Posselt | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-regression-tests/0fabb1c44f53911d0680c9d8743c5ec3c95dc52d/WCAG_2_2/1.7/PDFUA-Ref-2-10_Form.pdf) | `06d5cacd7c5ba9573a57c6d99d3f357ee14a24cff0c138036840fc44c6311806` | Adobe InDesign 15.0 |

## veraPDF corpus

Source: [veraPDF/veraPDF-corpus](https://github.com/veraPDF/veraPDF-corpus/tree/bb75f4f0073d9350dfd058c0162a367e6fadf25e), commit `bb75f4f0073d9350dfd058c0162a367e6fadf25e`.

License: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/), as stated in the repository's [README.md](https://github.com/veraPDF/veraPDF-corpus/blob/bb75f4f0073d9350dfd058c0162a367e6fadf25e/README.md).

Attribution: veraPDF test suite files © veraPDF Consortium, from https://github.com/veraPDF/veraPDF-corpus, licensed under CC BY 4.0. The files are unmodified.

| File | Source | SHA-256 | Producer |
| --- | --- | --- | --- |
| `verapdf-cidfont-type2.pdf` | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-corpus/bb75f4f0073d9350dfd058c0162a367e6fadf25e/PDF_A-2b/6.2%20Graphics/6.2.11%20Fonts/6.2.11.3%20Composite%20fonts/6.2.11.3.2%20CIDFonts/veraPDF%20test%20suite%206-2-11-3-2-t01-pass-a.pdf) | `ba939c67c486838167f13a9ef386a49d2edf8091494accf15824eadfaf40972e` | veraPDF test suite |
| `verapdf-cmap-embedded.pdf` | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-corpus/bb75f4f0073d9350dfd058c0162a367e6fadf25e/PDF_A-2b/6.2%20Graphics/6.2.11%20Fonts/6.2.11.3%20Composite%20fonts/6.2.11.3.3%20CMaps/veraPDF%20test%20suite%206-2-11-3-3-t01-pass-a.pdf) | `b322e2cb00f923d393e3018799338180ac0de135a3b6463326de8482203b3a33` | veraPDF test suite |
| `verapdf-composite-font.pdf` | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-corpus/bb75f4f0073d9350dfd058c0162a367e6fadf25e/PDF_A-2b/6.2%20Graphics/6.2.11%20Fonts/6.2.11.3%20Composite%20fonts/6.2.11.3.1%20General/veraPDF%20test%20suite%206-2-11-3-1-t01-pass-d.pdf) | `786d2056dbdee6f2c38334ed62738a414f550ae06558b79f6d9ea233901734fc` | veraPDF test suite |
| `verapdf-jpx-image.pdf` | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-corpus/bb75f4f0073d9350dfd058c0162a367e6fadf25e/PDF_A-2b/6.2%20Graphics/6.2.8%20Images/6.2.8.3%20JPEG2000/veraPDF%20test%20suite%206-2-8-3-t01-pass-a.pdf) | `f515bbe126ec675ea0b41b905507e8e638b65c8d5e1d99a15264e3217bf15076` | veraPDF Test Builder 1.0 |
| `verapdf-transparency-t05.pdf` | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-corpus/bb75f4f0073d9350dfd058c0162a367e6fadf25e/PDF_A-2b/6.2%20Graphics/6.2.10%20Transparency/veraPDF%20test%20suite%206-2-10-t05-pass-a.pdf) | `cb2859f9b779b07e0197ed4c3d33ba0ebfe65069631695b2563e3405d81bd40f` | veraPDF test suite |
| `verapdf-transparency-t07.pdf` | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-corpus/bb75f4f0073d9350dfd058c0162a367e6fadf25e/PDF_A-2b/6.2%20Graphics/6.2.10%20Transparency/veraPDF%20test%20suite%206-2-10-t07-pass-a.pdf) | `1e4036c679f3866a167ea7e0c99b1ed45bfa7a6a6cbe105046d8897ee378c841` | veraPDF test suite |
| `verapdf-optional-content.pdf` | [source](https://raw.githubusercontent.com/veraPDF/veraPDF-corpus/bb75f4f0073d9350dfd058c0162a367e6fadf25e/PDF_A-2b/6.9%20Optional%20content/veraPDF%20test%20suite%206-9-t03-pass-a.pdf) | `d186ec4942005768abc07e6d86669cf8ed10c0c979b1213824de2f6d0aa5fc9d` | veraPDF test suite |

## qpdf test suite

Source: [qpdf/qpdf](https://github.com/qpdf/qpdf/tree/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf), commit `4eba95899886e851cc41d76886483b347612f2a8`.

License: [Apache License 2.0](https://github.com/qpdf/qpdf/blob/4eba95899886e851cc41d76886483b347612f2a8/LICENSE.txt); notices in [NOTICE.md](https://github.com/qpdf/qpdf/blob/4eba95899886e851cc41d76886483b347612f2a8/NOTICE.md).

Notice: qpdf is copyright (c) 2005-2021 Jay Berkenbilt, 2022-2026 Jay Berkenbilt and Manfred Holger. Licensed under the Apache License, Version 2.0. The files are unmodified.

| File | Source | SHA-256 | Producer |
| --- | --- | --- | --- |
| `qpdf-linearized-lzw.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/lin5.pdf) | `4b716274089f173a1e2d826523e6e616182a1cafb267eff1e2bffa193440054b` | qpdf test suite |
| `qpdf-linearized.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/lin1.pdf) | `5bddf40fd6be12cadf33aca857a7dad2c58b9c3e4d3da0dc1676f44d71515543` | qpdf test suite |
| `qpdf-hybrid-xref.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/hybrid-xref.pdf) | `df1ae8b9cb2378aaeda6fdc11146dd37328737fd050941390d2c6bd188eea0e4` | qpdf test suite |
| `qpdf-incremental.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/incremental-1.pdf) | `5f9c43d3f201d8e9ea77aa162bf4eb527c563f9db3626a736455e003cf5baaf3` | qpdf test suite |
| `qpdf-enc-r3-rc4-user-view.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/enc-R3%2CV2%2CU%3Dview%2CO%3Dmaster.pdf) | `5c0683d687e2e020fe172b1dcb7c29fe94f67bc40f038fee2391de610053fc16` | qpdf test suite |
| `qpdf-enc-r6-aes256-user-view.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/enc-XI-R6%2CV5%2CU%3Dview%2CO%3Dmaster.pdf) | `96b9262bfacf1358c7a68ab46c11563677e858888950237fdac8f3a227210edf` | qpdf test suite |
| `qpdf-enc-r2-owner-only.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/enc-R2%2CV1%2CO%3Dmaster.pdf) | `4560d44658aca5b2b610f2c6631eb9db590fa43324c02890003dee6a8dec3b29` | qpdf test suite |
| `qpdf-enc-v4-aes.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/V4-aes.pdf) | `1d8a812844b8d8c7ba085a627eb3b7ab52aa651c6143ca07315084220a113240` | qpdf test suite |
| `qpdf-encrypted-images.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/encrypted-with-images.pdf) | `9f6c9ad93a462ad8a9c3c28debd5cf6ab80c60e7a8fe3b12e802b8e53aa0b1d4` | Apex PDFWriter (qpdf test suite) |
| `qpdf-damaged-stream.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/damaged-stream.pdf) | `91e119f2db2e9f738f696ddacea90498d460b0afda0fcdf71bb6281fc5085819` | qpdf test suite |
| `qpdf-append-xref-loop.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/append-xref-loop.pdf) | `7c67daccc433f574fbaf815d0d376cc79f40f95082ef1b231af436e9f7d27091` | qpdf test suite |
| `qpdf-bad-xref-entry.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/bad-xref-entry.pdf) | `1d3b6ed9362c92c3f24f6ee7650e1768387204b1bbb66327a4eba226ce801bb4` | qpdf test suite |
| `qpdf-recover-xref-stream.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/recover-xref-stream.pdf) | `9a78166f8d25c467ac45a39113728043c2a1595e74a39bb697d29565a0e5b2ac` | qpdf test suite |
| `qpdf-append-content-damaged.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/append-page-content-damaged.pdf) | `53ee40fa62175f0f7bf721497832ebaadf39b2551559ec28abdace73196f4e28` | qpdf test suite |
| `qpdf-deep-pages.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/deep-pages.pdf) | `1ee732a4e68421d5ae49aae1579304d7d9722b7a94fd1290ec3c55b42bf50ec7` | qpdf test suite |
| `qpdf-enc-20-pages-user-password.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/20-pages.pdf) | `4ea66d0e98525a5df9a766313e058bc3bfa2d28fe30a740d4cd8fe76d3dca391` | Apex PDFWriter (qpdf test suite) |
| `qpdf-inline-images.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/inline-images.pdf) | `8ae2fc89f4edd5579c2b05a2b9678388b29515143252ac40f6ec97eb7d97e1df` | GNU Ghostscript 5.50 (qpdf test suite) |
| `qpdf-form-from-odt.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/form-empty-from-odt.pdf) | `599304a4c47fa25155fc732cb9bffe2882ea39894e14f812905c4ec2f5826c14` | LibreOffice 6.0 Writer |
| `qpdf-form-filled-acrobat.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/form-filled-by-acrobat.pdf) | `465e30fe69d3b53d883ea1f01ea56a225ff0498f231ecda6f6fc986b9c5676ff` | LibreOffice 6.0 Writer, filled in Acrobat |
| `qpdf-form-utf8-appearances.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/need-appearances-utf8.pdf) | `aa4de49125baf4256a1fcb24dd0bdead029e7ff0363b086aef6774e76aae8b9f` | LibreOffice 6.1 Writer |
| `qpdf-rotated-annotations.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/rotated-shared-annotations-2.pdf) | `5fb0937663894589d4ed2dc5d3f2e64d971b836802f7f46f22a30be723b7f388` | qpdf test suite |
| `qpdf-utf16le.pdf` | [source](https://raw.githubusercontent.com/qpdf/qpdf/4eba95899886e851cc41d76886483b347612f2a8/qpdf/qtest/qpdf/utf16le.pdf) | `d7fa801d2bda9fe3b8c26003842ed8de4da21cbd3a768e8112002b4fe4fd0415` | qpdf test suite |

## ACL Anthology papers

Source: [ACL Anthology](https://aclanthology.org/).

License: [CC BY 4.0](https://creativecommons.org/licenses/by/4.0/), per the [ACL Anthology copyright policy](https://aclanthology.org/faq/copyright/).

Attribution: each paper below is © its authors, published by the Association for Computational Linguistics, and licensed under CC BY 4.0. The files are unmodified.

| File | Work | Source | SHA-256 | Producer |
| --- | --- | --- | --- | --- |
| `acl-2021-mt5-multilingual.pdf` | Linting Xue, Noah Constant, Adam Roberts, Mihir Kale, Rami Al-Rfou, Aditya Siddhant, Aditya Barua and Colin Raffel. 2021. "mT5: A Massively Multilingual Pre-trained Text-to-Text Transformer". NAACL-HLT 2021 | [source](https://aclanthology.org/2021.naacl-main.41.pdf) | `3416c10fe2483b147fe47f2eb9c7a87188ceb3f633db527d0c714cc640218b99` | pdfTeX (ACL NAACL 2021) |
| `acl-2023-arabicnlp-arabic.pdf` | Abdelrahman Mohamed, Fakhraddin Alwajih, El Moatez Billah Nagoudi, Alcides Alcoba Inciarte and Muhammad Abdul-Mageed. 2023. "Violet: A Vision-Language Model for Arabic Image Captioning with Gemini Decoder". ArabicNLP 2023 | [source](https://aclanthology.org/2023.arabicnlp-1.1.pdf) | `d9415ee78237939ab928afa7d4fbcc7af52ce13038db9d80ad04ad611d1bfef5` | pdfTeX 1.40.24 (ArabicNLP 2023) |
| `acl-2021-wanlp-arabic.pdf` | Ahmed Abdelali, Hamdy Mubarak, Younes Samih, Sabit Hassan and Kareem Darwish. 2021. "QADI: Arabic Dialect Identification in the Wild". WANLP 2021 | [source](https://aclanthology.org/2021.wanlp-1.1.pdf) | `84c01926bee185aca0e22aa05f09ecbccb823f17401124f2b485d2cca606b8cd` | pdfTeX 1.40.20 (WANLP 2021) |

## United States government works

Sources: the National Institute of Standards and Technology and the National Archives and Records Administration.

License: works of the United States government are in the public domain in the United States ([17 U.S.C. § 105](https://www.usa.gov/government-copyright)). No attribution is required; the source is recorded for provenance.

| File | Source | SHA-256 | Producer |
| --- | --- | --- | --- |
| `nist-sp800-63-3.pdf` | [source](https://nvlpubs.nist.gov/nistpubs/SpecialPublications/NIST.SP.800-63-3.pdf) | `915bb812cb705918461e642e84141c6ca2adc2ca0106e1d88f3884c61f07110a` | Microsoft Word via Acrobat PDFMaker 17 |
| `nist-cswp-29.pdf` | [source](https://nvlpubs.nist.gov/nistpubs/CSWP/NIST.CSWP.29.pdf) | `3c31f46fee98cac0c4323453e5109291a213b4de7fef8c058af9bf67f717433c` | Microsoft Word via Acrobat PDFMaker 23 |
| `nara-jfk-104-10003-10041.pdf` | [source](https://www.archives.gov/files/research/jfk/releases/2025/0318/104-10003-10041.pdf) | `c3c7e7b44e6f65e08d4e73950b74428f40d26e7ead49e945f003f6b8f0903e48` | Aspose.PDF for .NET 23.2.0 scan |

## RFC Editor

Source: the [RFC Editor](https://www.rfc-editor.org/).

License: [IETF Trust Legal Provisions](https://trustee.ietf.org/assets/ietf-trust-legal-provisions-tlp/), which allow RFCs to be copied and distributed in full without modification.

Notice: RFC 791, "Internet Protocol", J. Postel (editor), September 1981. The file is an unmodified copy of the RFC Editor's PDF and keeps any notice it carries.

| File | Source | SHA-256 | Producer |
| --- | --- | --- | --- |
| `rfc791-ghostscript.pdf` | [source](https://www.rfc-editor.org/rfc/pdfrfc/rfc791.txt.pdf) | `28f377bf7e65440818f4c8d8e1e65a48bb6c3e41f118911064098367aa039fc4` | Aladdin Ghostscript 5.10 |

## Japanese Digital Agency

Source: the [Digital Agency (デジタル庁) website](https://www.digital.go.jp/).

License: [Public Data License 1.0 (公共データ利用規約 第1.0版)](https://www.digital.go.jp/en/copyright-policy), compatible with CC BY 4.0.

Attribution: 出典：デジタル庁ウェブサイト (Source: Digital Agency website), at the URL listed for each file. The files are unmodified.

| File | Work | Source | SHA-256 | Producer |
| --- | --- | --- | --- | --- |
| `jp-digital-agency-word-betten.pdf` | 「政府標準利用規約（第2.0版）」の解説, 内閣官房IT総合戦略室 (Cabinet Secretariat IT Strategy Office) | [source](https://www.digital.go.jp/assets/contents/node/basic_page/field_ref_resources/f7fde41d-ffca-4b2a-9b25-94b8a701a037/a0f187e6/20220706_resources_data_betten_01.pdf) | `7aa72a369c0e684bf0f9253485ca27099d5032a64b97d3b954cbb0a31ba834c4` | Microsoft Word (Japanese) via Acrobat PDFMaker 11 |
| `jp-digital-agency-justpdf-outline.pdf` | 「公共データ利用規約（第1.0版）」の解説, デジタル庁 (Digital Agency) | [source](https://digital.go.jp/assets/contents/node/basic_page/field_ref_resources/f7fde41d-ffca-4b2a-9b25-94b8a701a037/b44a7e0c/20240705_resources_data_outline_07.pdf) | `7b6157cb8a0536b36bc44ede57741056cedbd2b1d70303aa5a1f06f6dbfb2a56` | JUST PDF 4 (Ichitaro) |
