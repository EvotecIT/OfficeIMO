# OfficeIMO.Email

`OfficeIMO.Email` is the dependency-free OfficeIMO engine for offline email and Outlook artifacts.

The package owns bounded reading, structured diagnostics, and deterministic writing for EML/MIME and Outlook MSG. TNEF (`winmail.dat`), mbox archives, and typed Outlook item projections use the same model as they are added. Network protocols, authentication, DKIM, ARC, PGP, and S/MIME cryptographic operations remain outside this package.

```csharp
var reader = new EmailDocumentReader();
EmailReadResult result = reader.Read("message.eml");

var writer = new EmailDocumentWriter();
writer.Write(result.Document, "copy.eml", EmailFileFormat.Eml);
```

Signed or encrypted MIME entities are preserved as content. `OfficeIMO.Email` does not verify signatures, resolve keys, or decrypt payloads.
