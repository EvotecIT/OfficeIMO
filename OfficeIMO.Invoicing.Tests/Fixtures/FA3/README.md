# Official Polish FA(3) examples

These 26 XML files are unchanged entries from the Ministry of Finance
[FA(3) example archive](https://ksef.podatki.gov.pl/media/e5cia0ey/przykladowe-pliki-dla-struktury-logicznej-e-faktury-fa-3.zip).
The local ASCII filenames keep their original example numbers. `manifest.json`
records the original entry names, archive SHA-256 and exact file SHA-256 values.

The fixtures contain the authority's fictional example identities. They cover
ordinary VAT invoices, signed corrections, advances, settlements, simplified
invoices, foreign currency and settlement adjustments. Core tests verify bounded
reading, exact source preservation and explicit projection limits. Validation
tests link these same files and qualify them against the pinned official XSD.
Schema success is independent of common-model completeness and fiscal acceptance.
