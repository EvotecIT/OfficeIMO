# Security and protected-content capabilities

OfficeIMO keeps package structure and cryptography separate. Document packages can inspect signature carriers and enforce safe mutation rules without `OfficeIMO.Security`. Applications that create or cryptographically validate signatures pass an `IOfficeSecurityProvider` explicitly; installing a document package does not install `OfficeIMO.Security` transitively.

## Capability inventory

| Package | Structural inspection | Mutation safety | Password protection | Cryptographic signing and validation | Encryption and decryption | Optional provider | Deliberately unsupported |
| --- | --- | --- | --- | --- | --- | --- | --- |
| `OfficeIMO.Word` | OPC and VBA signature carriers; legacy, agile, and V3 VBA profile evidence | Signed-package edits block by default; explicit remove or preservation policies | OOXML editing restrictions; password-to-open load and save | Cross-platform OPC and managed VBA creation, content binding, and validation | Office password-to-open encryption/decryption | Required for signature creation, CMS/XML validation, trust, revocation, and timestamps | VBA source execution/editing; password/key recovery |
| `OfficeIMO.Excel` | OPC and VBA signature carriers, including XLSB; legacy, agile, and V3 VBA profile evidence | Signed-workbook edits block or require an explicit invalidation policy | Workbook/worksheet protection; password-to-open load and save for Open XML and supported legacy XLS profiles | Cross-platform OPC and managed VBA creation, content binding, and validation, including XLSB | Office password-to-open encryption/decryption | Required for signature creation and cryptographic validation | VBA source execution/editing; compound-file XLS signature creation |
| `OfficeIMO.PowerPoint` | OPC, legacy binary signature metadata, and VBA carriers; legacy, agile, and V3 VBA profile evidence | Signed-presentation edits block or require an explicit invalidation policy | Password-to-open load and save for Open XML and binary PPT | Cross-platform OPC and managed VBA creation, content binding, and validation | Open XML password encryption plus legacy PPT RC4 CryptoAPI interoperability | Required for signature creation and cryptographic validation | VBA source execution/editing; compound-file PPT signature creation or cryptographic validation |
| `OfficeIMO.Visio` | OPC signature origin, relationships, and signature parts | Loaded signed diagrams block rebuilding by default | None | OPC create/validate | None | Required for signature creation and cryptographic validation | VSD/VSDM VBA signing; password encryption/decryption |
| `OfficeIMO.OpenDocument` | Signature carriers and encrypted-manifest declarations | Unchanged signatures are preserved; changed signed packages block unless stale signatures are explicitly removed; encrypted sources cannot be saved as plaintext accidentally | ODT/ODS/ODP password-to-open load and save | Creates and validates the bounded OfficeIMO XML package-manifest signature profile in `META-INF/documentsignatures.xml` | AES-256-CBC encryption/decryption with bounded PBKDF2 and classified failures | Required only for XML signature creation and cryptographic validation | Legacy Blowfish encryption, password recovery, and general third-party ODF signature profiles |
| `OfficeIMO.Epub` | `META-INF/signatures.xml` plus encryption and obfuscation declarations | Reader APIs preserve the source artifact; package signing commits atomically | None | Creates and validates the bounded OfficeIMO XML package-manifest signature profile in `META-INF/signatures.xml` | IDPF and Adobe font deobfuscation when package identity supplies the key; no DRM decryption | Required for XML signature creation and cryptographic validation | DRM removal, password recovery, arbitrary XML Encryption, and general third-party EPUB signature profiles |
| `OfficeIMO.Pdf` | Signature dictionaries, byte ranges, revisions, and encryption metadata | Full rewrites of signed PDFs block; supported incremental updates preserve prior revisions | PDF Standard user/owner passwords and permission bits | External-signing contracts are first-party; built-in CMS creation/validation is provider-backed | PDF Standard security read, write, remove, and re-encrypt workflows | Required only for built-in CMS, timestamp, and X.509 operations | Password recovery and unsupported security handlers or crypt filters |
| `OfficeIMO.Email` | Clear/opaque S/MIME and OpenPGP/MIME wrapper detection with original protected-artifact retention | Protected projections are returned separately; unchanged protected source bytes remain available | None | Clear and opaque S/MIME creation and verification | Caller-selected CMS EnvelopedData recipient encryption/decryption and sign-then-encrypt | Required for S/MIME CMS/X.509 operations | Certificate discovery, DKIM, ARC, OpenPGP cryptography, and password/key recovery |
| `OfficeIMO.Core` | Owns the bounded shared OPC, VBA, and XML-package signature structures | Provides atomic file-commit and resource-limit primitives | None | Defines `IOfficeSecurityProvider`; contains no concrete cryptographic provider | None | Concrete implementation is supplied by `OfficeIMO.Security` or an application provider | Certificate/private-key discovery or storage |
| `OfficeIMO.Security` | No document parser | Not applicable | None | CMS, XML Signature, X.509 trust/revocation, and timestamp services | CMS EnvelopedData encryption/decryption used by S/MIME | This is the optional concrete provider package | Document-format mutation policy and password recovery |

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

Macro-capable Word, Excel, and PowerPoint packages share managed MS-OVBA canonicalization, legacy/agile/V3 profile writers, provider-backed CMS, and content-binding validation on every supported platform. Supported ZIP package extensions are `.docm`, `.dotm`, `.xlsm`, `.xltm`, `.xlam`, `.xlsb`, `.pptm`, `.potm`, `.ppsm`, and `.ppam`. Producer-corpus hashes cover all three profiles across DOCM, XLSM, XLSB, and PPTM. A registered Microsoft Office SIP can be enabled on Windows as an additional differential check for the legacy and V3 transcripts; its public indirect-data API does not expose an unambiguous agile-only digest, so agile validation remains managed and corpus-bound. The SIP is never a runtime prerequisite. Create VBA signatures first, then the OPC package signature because changing VBA signature relationships invalidates an existing OPC signature.

OpenDocument and EPUB expose the same bounded XML package-manifest profile through their host APIs:

```csharp
OdfDocument.SignPackage("report.odt", security, signingCertificate);
EpubDocument.SignPackage("book.epub", security, signingCertificate);
```

The created XML Signature authenticates an exact, bounded manifest of ZIP entry paths and SHA-256 digests. Validation rejects missing, changed, duplicate, or unsigned entries. It does not claim compatibility with every producer-specific ODF or EPUB signature profile.

## ODF password encryption

`OdfLoadOptions.Password` opens ODT, ODS, and ODP packages using the interoperable AES-256-CBC profile with a SHA-256 password start key, per-entry PBKDF2-HMAC-SHA1, and SHA-256/1K checksums. The accepted PBKDF2 policy is 10,000 through 10,000,000 iterations; output defaults to 100,000. `OdfSaveOptions.Encryption` creates the same profile with fresh salt and initialization-vector material for every encrypted entry. Passwords are UTF-8, operation-scoped, and never retained by the document model.

Missing passwords, wrong passwords, unsupported profiles, malformed metadata, and resource-limit failures are classified through `OdfEncryptedPackageException`. A document loaded from encrypted source requires new encryption settings at save time; plaintext output requires an explicit `OdfEncryptionHandling.Remove` decision. The corpus includes hash-pinned LibreOffice output, and interoperability evidence covers OfficeIMO reading LibreOffice output and LibreOffice reading OfficeIMO output. Legacy Blowfish profiles and password recovery remain unsupported.

## Machine-readable coverage

[`OfficeProtectionCapabilityCatalog.Current`](Compatibility/generated/protected-content.md) is the generated operation-level contract. It keeps password encryption, recipient encryption, digital signatures, reversible font obfuscation, editing restrictions, and checksum-based access deterrence as separate mechanism types.
