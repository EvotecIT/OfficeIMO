# OfficeIMO.Email.Data

`OfficeIMO.Email.Data` is the small convenience API included in `OfficeIMO.Email` for applications that receive mixed
persisted email and Outlook data. `EmailDataArtifact.Open` detects the source and returns the existing email, Store,
or AddressBook public type.

```csharp
using OfficeIMO.Email.Data;

using EmailDataOpenResult result = EmailDataArtifact.Open(path);

switch (result.Kind) {
    case EmailDataArtifactKind.EmailDocument:
        Console.WriteLine(result.EmailDocument!.Subject);
        break;
    case EmailDataArtifactKind.Store:
        Console.WriteLine(result.Store!.Folders.Count);
        break;
    case EmailDataArtifactKind.OfflineAddressBook:
        Console.WriteLine(result.AddressBook!.DeclaredEntryCount);
        break;
}
```

The facade does not copy parsers, property models, store queries, or mutation logic. Advanced callers can use the
owning API directly. For ambiguous paths, set `EmailDataOpenOptions.ExpectedKind`. Full email reads remain
the default so protected-artifact pass-through is preserved; opt into `UseStreamingEmailReader` when file-backed
attachment content is preferred.

Profile discovery, account settings, synchronization state, search indexes, `.nk2`, and modern Outlook autocomplete
caches are intentionally outside this API. They need a real consumer and format-specific safety contract before a
separate Profile owner is justified.
