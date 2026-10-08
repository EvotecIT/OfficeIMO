# SECTIONPAGES native fixture import boundary

The four Microsoft Word desktop 16.0 binary DOC fixtures in this folder retain
SECTIONPAGES instructions as editable fields and retain their cached result of
`1`. Header/footer table structure, visible numbering and section boundaries
are separate observable contracts. These fixtures establish field import;
they do not claim exact pagination or preservation of arbitrary binary payloads.

Each fixture reports `DOC-BINARY-DATA-STREAM-PRESENT` for a 4,096-byte `Data`
stream. The source retains that payload; the projected document does not carry
it as editable content. Each also reports `DOC-QUICK-SAVE-HISTORY-PRESENT` for
15 quick saves. The readable content imports, while quick-save history is not
projected as editable revision history. No unsupported readable feature or
import error is reported for these fixtures.

The field test performs read-only import and inspection. The ordinary loss
preflight still applies before converting or saving a projected document;
this report does not waive that gate.
