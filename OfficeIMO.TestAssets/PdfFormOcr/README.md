# Form OCR fixtures

The scanned and rotated synthetic documents were produced independently with ReportLab 4.4.9 and pypdf 6.10.0. Poppler 26.09.0 rendered the native filled form into scan images. Transparent empty AcroForm widgets retain the field definitions over those images. Values read from the PDF remain empty until a reviewed proposal is applied; the read-only field retains `KEEP`.

The second fixture adds a CropBox, 90-degree rotation and UserUnit 2 on page two. The Amount field uses separated parent-field numeric helper actions. Country has distinct display and export values. CustomRule contains unsupported script code.

The certified fixture retains the original ReportLab filled values and appearances and adds synthetic DocMDP permission level 2 metadata. It qualifies append-only form updates and refreshed appearances while preserving the original byte prefix. Its signature bytes do not form a valid cryptographic signature.

`manifest.json` records expected geometry, values, producer versions and SHA-256 hashes. `producer.py <output-directory>` regenerates candidate fixtures into an explicit disposable directory. The producer requires validation-only ReportLab, pypdf and Poppler; none is an OfficeIMO runtime dependency. ReportLab and pypdf use BSD licenses; Poppler is GPL-licensed tooling. The generated fixture content and this producer script are covered by the repository license.

These fixtures qualify form parsing, geometry, reviewed value handling and saved appearances. They do not establish general OCR accuracy or real signature acceptance.
