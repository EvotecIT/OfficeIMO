# PDF editing fixtures

`source-font-reportlab.pdf` contains a conventional horizontal TrueType text run
with an embedded subset and a ToUnicode map. Its Polish characters distinguish
Unicode mapping from byte-value extraction. The encrypted copies add independent
Standard-security output through pypdf: 40-bit RC4, AES-128 and AES-256.

`manifest.json` records producer versions, source text, test passwords and hashes.
The passwords are fixture data. The font comes from the existing baseline font
assets and retains their license and attribution.

Run `generate_fixtures.py` manually with ReportLab and pypdf to regenerate the
corpus. Those packages are optional validation tools; ordinary builds and tests
read the checked-in fixtures without requiring Python. ReportLab uses invariant
output, while encrypted fixture regeneration changes random cryptographic data
and therefore requires updating the manifest.
