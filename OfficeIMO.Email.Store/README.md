# OfficeIMO.Email.Store

`OfficeIMO.Email.Store` reads mailbox and email-store formats into the common `OfficeIMO.Email.EmailDocument` model.
It is fully managed and does not require Outlook, native libraries, or third-party parser packages at runtime.

The public API is bounded by default:

```csharp
EmailStoreReadResult result = new EmailStoreReader().Read("archive.pst");

foreach (EmailStoreFolder folder in result.Store.Folders) {
    foreach (EmailStoreMessage message in folder.Messages) {
        Console.WriteLine(message.Document.Subject);
    }
}
```

PST and OST files are read-only. The reader projects their MAPI properties through `OfficeIMO.Email`, so MSG,
OFT, PST, and OST items share the same message, appointment, contact, task, journal, and note semantics.
