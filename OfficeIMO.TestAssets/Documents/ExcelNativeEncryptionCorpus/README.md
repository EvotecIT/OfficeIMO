# Native XLS encryption corpus

Microsoft Excel for Mac 16.113.4 produced the plain and CryptoAPI-encrypted XLS
controls in this folder. Each contains one embedded XLSX package, numeric values,
a formula and document properties. The embedded source is an owned, macro-free
OfficeIMO workbook. [provenance.json](provenance.json) records hashes, passwords,
native observations and limits.

The encrypted input has FilePass version 4.2 and flags `0x0c`. Its `encryption`
stream contains a compound-storage descriptor, without encrypted document-property
descriptors. Excel also retains a clear copy of that storage. The copies have the
same hierarchy, class/state and payload bytes, with different root timestamps.
OfficeIMO accepts matching copies and rejects conflicting payloads or identities.

OfficeIMO edits A2 to `Edited native embedding`, changes B1 to `40+2`, and saves
plain and password-rotated XLS. Excel opens both exact outputs without repair
prompts. After F9 and SaveAs, the checked-in readbacks retain A2, B1=42, B3=50,
the title/subject and the embedded workbook's marker and value 42. This proves
the named package-preservation cohort; it does not qualify OLE activation,
external links, ActiveX, every embedded server or rendered geometry.

## Regenerate the Mac controls

Copy `embedded-source.xlsx` and the reviewed
[qualification procedure](../../../Build/Verification/ExcelLegacyEmbeddingOracle.bas)
beside an owned macro-enabled controller. Import the procedure into its standard
module and explicitly run `QualifyNativeEmbedding` in desktop Excel. It creates
a separate workbook, inserts the package, saves plain and password-protected XLS,
and closes that workbook. It has no automatic entry points and changes no global
application security settings. Do not run macros from untrusted fixture files.

The input password is `test-password`; the encrypted readback uses `new-password`.
Regenerate the edited controls through the public `LoadEncrypted`, `CellValue`,
`CellFormula`, `Save` and `SaveEncrypted` APIs, then open the exact outputs in
Excel, recalculate and save separate readbacks. Check their values, properties
and embedded payload through the public API before updating provenance.

## Remaining Windows encrypted-summary check

Mac Excel returns error 5 for both the explicit CryptoAPI provider call and the
properties-only option call. Its resulting property streams are plaintext, so
these files do not qualify a native producer of encrypted summary properties.
The following is a pending procedure for an existing Windows desktop Excel
session, not a completed Windows result:

1. Use an owned workbook with the same title, subject, cells and embedded package.
   Record the Windows and Excel versions. Preserve existing user work and security
   settings; use the reviewed controller rather than enabling fixture macros.
2. Before saving XLS, call the native object model with all arguments:

   ```vb
   wb.SetPasswordEncryptionOptions _
       PasswordEncryptionProvider:="Microsoft Enhanced Cryptographic Provider v1.0", _
       PasswordEncryptionAlgorithm:="RC4", _
       PasswordEncryptionKeyLength:=128, _
       PasswordEncryptionFileProperties:=True
   wb.SaveAs Filename:=ownedOutputPath, FileFormat:=56, Password:="test-password"
   ```

3. Inspect the actual FilePass header and decrypted `encryption` descriptor list.
   Require the document-property encryption bit to indicate encrypted properties
   and both `\u0005SummaryInformation` and `\u0005DocumentSummaryInformation`
   descriptors, alongside the embedded storage. A password prompt or a successful
   SaveAs alone does not establish that profile. If the native call fails or
   properties remain plaintext, record that negative result and the exact build.
4. Load that native input through `ExcelDocument.LoadEncrypted`; assert exact
   title/subject, cell values/formula and embedded package. Edit a cell/formula,
   save plain XLS and password-rotated XLS, and reopen both in Windows Excel
   without repair. Recalculate, save separate readbacks and verify the same
   properties, values and embedded content through OfficeIMO. Retain the input,
   readbacks, hashes and producer evidence with the corpus.

No Windows producer result is claimed until these acceptance checks pass.
