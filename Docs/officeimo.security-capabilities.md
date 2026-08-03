# Security and protected-content capabilities

OfficeIMO keeps package structure and cryptography separate. Document packages can inspect signature carriers and enforce safe mutation rules without `OfficeIMO.Security`. Applications that create or cryptographically validate signatures pass an `IOfficeSecurityProvider` explicitly; installing a document package does not install `OfficeIMO.Security` transitively.

## Capability inventory

| Package | Structural inspection | Mutation safety | Password protection | Cryptographic signing and validation | Encryption and decryption | Optional provider | Deliberately unsupported |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `OfficeIMO.Word` | OPC and VBA signature carriers; legacy, agile, and V3 VBA profile evidence | Signed-package edits block by default; explicit remove or preservation policies | OOXML editing restrictions; password-to-open load and save | OPC create/validate on supported platforms; VBA create and Office SIP content-binding validation on Windows | Office password-to-open encryption/decryption | Required for signature creation, CMS/XML validation, trust, revocation, and timestamps | VBA source execution/editing; VBA content-binding claims without the registered Microsoft Office SIP |
| `OfficeIMO.Excel` | OPC and VBA signature carriers, including XLSB; legacy, agile, and V3 VBA profile evidence | Signed-workbook edits block or require an explicit invalidation policy | Workbook/worksheet protection; password-to-open load and save for Open XML and supported legacy XLS profiles | OPC create/validate; VBA create and Office SIP content-binding validation on Windows | Office password-to-open encryption/decryption | Required for signature creation and cryptographic validation | VBA source execution/editing; compound-file XLS signature creation |
| `OfficeIMO.PowerPoint` | OPC, legacy binary signature metadata, and VBA carriers; legacy, agile, and V3 VBA profile evidence | Signed-presentation edits block or require an explicit invalidation policy | Password-to-open load and save for Open XML and binary PPT | OPC create/validate; VBA create and Office SIP content-binding validation on Windows | Open XML password encryption plus legacy PPT RC4 CryptoAPI interoperability | Required for signature creation and cryptographic validation | VBA source execution/editing; compound-file PPT signature creation or cryptographic validation |
| `OfficeIMO.Visio` | OPC signature origin, relationships, and signature parts | Loaded signed diagrams block rebuilding by default | None | OPC create/validate | None | Required for signature creation and cryptographic validation | VSD/VSDM VBA signing; password encryption/decryption |
| `OfficeIMO.OpenDocument` | Signature carriers and encrypted-manifest declarations | Unchanged signatures are preserved; changed signed packages block unless stale signatures are explicitly removed | None | Creates and validates the bounded OfficeIMO XML package-manifest signature profile in `META-INF/documentsignatures.xml` | None; encrypted packages are detected and rejected before editing | Required for XML signature creation and cryptographic validation | General third-party ODF signature profiles; ODF encryption/decryption until an interoperable producer corpus and explicit password/key policy exist |
| `OfficeIMO.Epub` | `META-INF/signatures.xml` plus encryption and obfuscation declarations | Reader APIs preserve the source artifact; package signing commits atomically | None | Creates and validates the bounded OfficeIMO XML package-manifest signature profile in `META-INF/signatures.xml` | No DRM or resource decryption | Required for XML signature creation and cryptographic validation | DRM removal, password recovery, arbitrary XML Encryption, and general third-party EPUB signature profiles |
| `OfficeIMO.Pdf` | Signature dictionaries, byte ranges, revisions, and encryption metadata | Full rewrites of signed PDFs block; supported incremental updates preserve prior revisions | PDF Standard user/owner passwords and permission bits | External-signing contracts are first-party; built-in CMS creation/validation is provider-backed | PDF Standard security read, write, remove, and re-encrypt workflows | Required only for built-in CMS, timestamp, and X.509 operations | Password recovery and unsupported security handlers or crypt filters |
| `OfficeIMO.Email` | Clear/opaque S/MIME and OpenPGP/MIME wrapper detection with original protected-artifact retention | Decryption returns separate content and retains the source protected message | None | S/MIME verification | Caller-selected CMS EnvelopedData recipient decryption | Required for S/MIME verification and decryption | S/MIME creation, certificate discovery, DKIM, ARC, OpenPGP cryptography, and password/key recovery |
| `OfficeIMO.Drawing` | Owns the bounded shared OPC, VBA, and XML-package signature structures | Provides atomic file-commit and resource-limit primitives | None | Defines `IOfficeSecurityProvider`; contains no concrete cryptographic provider | None | Concrete implementation is supplied by `OfficeIMO.Security` or an application provider | Certificate/private-key discovery or storage |
| `OfficeIMO.Security` | No document parser | Not applicable | None | CMS, XML Signature, X.509 trust/revocation, and timestamp services | CMS EnvelopedData decryption used by S/MIME | This is the optional concrete provider package | Document-format mutation policy and password recovery |

“Password protection” and “encryption” are separate. Word editing restrictions and Excel workbook/worksheet protection discourage modification but do not encrypt file contents. Password-to-open and PDF Standard security do encrypt content.

## Shared signature APIs

Word, Excel, PowerPoint, and Visio use the same bounded OPC implementation:

```csharp
using OfficeIMO.Excel;
using OfficeIMO.Security;

IOfficeSecurityProvider security = OfficeSecurityProvider.Default;
ExcelDocument.SignPackageSignature("report.xlsx", security, signingCertificate);

var options = new OfficePackageSignatureValidationOptions {
    ValidateCertificateTrust = true
};
OfficePackageSignatureValidationReport report =
    ExcelDocument.ValidatePackageSignatures("report.xlsx", security, options);
```

`InspectPackageSignatures(...)`, `ValidatePackageSignatures(...)`, `SignPackageSignature(...)`, and
`TrySignPackageSignature(...)` are the cross-host contract. Word's established `InspectSignatures()` API projects the
same shared structural result into its compatibility model; its established `ValidateSignatures(...)` API remains an
additional Word-specific timestamp and diagnostic surface over the shared archive, transform, and writer primitives.

Macro-capable Word, Excel, and PowerPoint packages share the bounded VBA inspector and Windows Office SIP workflow. Supported ZIP package extensions are `.docm`, `.dotm`, `.xlsm`, `.xltm`, `.xlam`, `.xlsb`, `.pptm`, `.potm`, `.ppsm`, and `.ppam`. Create the VBA signatures first, then the OPC package signature; changing VBA signature relationships invalidates an existing OPC signature.

OpenDocument and EPUB expose the same bounded XML package-manifest profile through their host APIs:

```csharp
OdfDocument.SignPackage("report.odt", security, signingCertificate);
EpubDocument.SignPackage("book.epub", security, signingCertificate);
```

The created XML Signature authenticates an exact, bounded manifest of ZIP entry paths and SHA-256 digests. Validation rejects missing, changed, duplicate, or unsigned entries. It does not claim compatibility with every producer-specific ODF or EPUB signature profile.

## ODF encryption boundary

OfficeIMO does not create or decrypt encrypted ODF packages. The loader detects manifest encryption before exposing an editable model and throws `OdfEncryptedPackageException`; unchanged signed package content remains preservation-aware. Encryption will not be added from specification examples alone: it requires a producer-identified interoperability corpus, an explicit password/key derivation policy, wrong-key and partial-failure tests, and proof that failed operations leave the original package unchanged.
