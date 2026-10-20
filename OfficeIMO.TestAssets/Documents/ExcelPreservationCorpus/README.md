# Native workbook preservation controls

The XLSM control contains a new inert VBA project: host identity attributes and
comments, with no procedures or event handlers. Its worksheet/chart content is
an OfficeIMO projection of the repository's CC0 OpenPreserve `valid.xls` fixture.
The original VBA was replaced. Native Excel opened the corrected XLSM without
repair and saved the paired XLSB and XLS files with macros disabled.

`provenance.json` records exact hashes, source attribution and the native encoder.
These files qualify a preservation cohort; they do not establish every legacy
workbook feature, chart appearance, or macro execution behavior. Fixture tools
are independent validation and are not product runtime requirements.

The populated control adds a constant bar series to `b_chart`: Alpha 10, Beta 20,
and Gamma 30. Native Excel displays these bars and encodes the paired
`excel-populated-macro.xlsb` and `.xls` files. This makes visible chart
preservation testable; the other chart sheets remain empty controls.

`officeimo-populated-edited.xlsb` changes only `graphs2!A1`. All ZIP entries
except the first worksheet binary part match the native source byte-for-byte.
Excel reopened it without repair, displayed the edited text and the same populated
bar chart, and kept macros disabled. VBA execution remains outside this proof.

The plain and encrypted edited XLS controls use the same existing `graphs2!A1`
literal edit. Native Excel reopened the exact read-only artifacts without repair,
with macros disabled, and displayed the edited cell and populated chart. Their
hashes remained unchanged after inspection. The encrypted control uses the public
test password `test-password`; it contains the same inert VBA project. These
controls cover preservation during supported edits, not chart editing or macro
execution. See `provenance.json` for artifact hashes and acceptance details.

The `officeimo-lifecycle` controls exercise worksheet rename, add, remove and
reorder while retaining the populated chart and inert VBA hosts. The encrypted
controls use `test-password`. Native Excel accepts all four without repair and
recalculates the added sheet's cross-sheet formula to 43. The paired
`excel-lifecycle-readback` files were saved by Excel before tab navigation, so
their tab order and VBA identities can be checked independently.

The `excel-array-readback` files were saved by native Excel after reopening an
OfficeIMO two-cell array formula, styling its follower and clearing its caches.
Excel recognizes both array members and recalculates them to 5. These fixtures
protect native group import and cache/identity contracts; they do not establish
unrestricted array calculation. Exact procedures and hashes are in
`provenance.json`.
