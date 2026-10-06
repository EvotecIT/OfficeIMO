# OfficeIMO.Email.Html

## Export a local HTML copy

```csharp
using OfficeIMO.Email;

OfficeIMO.Email.EmailMessage message = EmailMessageReader.Read("message.msg").Messages[0];
EmailHtmlExportResult copy = EmailHtmlExporter.Export(message, "message.html", includeAttachments: true);
Console.WriteLine(copy.Path);
Console.WriteLine(copy.AssetsDirectory);
```

The exporter selects HTML, then RTF, then plain text, removes active content, and blocks remote resources. Embedded attachments are extracted into a fresh companion resource directory, and supported image attributes, responsive sources, and CSS URLs are rewritten through the shared HTML resource policy. `includeAttachments` also saves regular files and adds download links. Extraction results retain original names, logical attachment indexes, byte counts, and hashes.

Keep the HTML file and its `.assets-<id>` directory together. HTML commits use an atomic fail-or-replace policy. An overwrite creates a fresh resource directory; earlier resource files survive cancellation or later failures and may be removed when no longer needed. Native email artifacts remain the appropriate choice when original message structure and format-specific metadata matter.

`OfficeIMO.Email.Html` is the optional HTML bridge for `OfficeIMO.Email`. It selects an HTML, RTF, or plain-text body once, applies the shared untrusted HTML policy, indexes CID and content-location resources, and exposes a prepared `HtmlConversionDocument` to renderers and readers.

```powershell
dotnet add package OfficeIMO.Email.Html
```

```csharp
using OfficeIMO.Email;

EmailBodyProjectionResult projection = EmailBodyProjection.Create(message);
string safeHtml = projection.Html;
EmailBodyResource? logo = projection.ResolveResource("cid:logo@example.test");

if (logo is not null) {
    using FileStream output = File.Create("logo.bin");
    await logo.CopyToAsync(output);
}
```

For targeted image inspection and rewriting without creating a full body projection, use the bounded, network-free DOM view:

```csharp
EmailHtmlImageDocument images = EmailHtmlImageDocument.Parse(message.Body.Html ?? string.Empty);

foreach (EmailHtmlImageReference image in images.Images) {
    if (image.Source == "logo.png") {
        images.SetImageSource(image.Index, "cid:logo@example.test");
    }
}

string rewrittenHtml = images.ToHtml();
```

`Parse` accepts either an HTML fragment or a complete document. `Images` reports explicit `img[src]` attributes in document order, `SetImageSource` rewrites one stable index, and `ToHtml` preserves fragment-versus-document shape. This workflow does not open local files or download remote resources.

Remote resources are blocked by default. Selecting `AllowByConsumerResolver` only retains eligible HTTP(S) references; this package never downloads them. Attachment content remains operation-scoped and is opened only through bounded resource reads.

Set `IncludeResourceReferences` to `false` to exclude resource URLs from projected markup through the shared URL policy. `MaxBodySourceCharacters` optionally checks the selected original body before HTML encoding, RTF reading or plain-text fallback; shared HTML limits separately bound the generated projection.

`EmailBodyProjectionOptions` bounds each projection by indexed resource count, bytes per resource, and declared or read bytes across all resources. The defaults are 128 resources, 128 MiB per resource, and 256 MiB per projection. `OpenReadStream`, `CopyTo`, and their asynchronous counterparts let consumers process content without first allocating another full byte array. Repeated reads share the projection-wide budget. Body-only consumers can set `IncludeResources` to `false` to avoid indexing attachments.

`OfficeIMO.Email.Image` uses the prepared HTML and resource index for rendering. `OfficeIMO.Reader.Email` uses the same projection before producing safe text or Markdown, so those adapters do not choose bodies, sanitize markup, or resolve embedded resources independently.

The core `OfficeIMO.Email` package does not depend on HTML libraries. Install this bridge only when safe HTML, RTF fallback, resource resolution, rendering, or Markdown projection is needed. The dependency footprint is `OfficeIMO.Email`, `OfficeIMO.Html`, and `OfficeIMO.Html.Rtf`.

For indexing, `EmailIndexText` returns markup-free text and recognized quote/signature regions:

```csharp
EmailIndexTextResult index = EmailIndexText.Create(message, new EmailIndexTextOptions {
    ExcludeQuotes = true,
    ExcludeSignatures = true
});
string textForSearch = index.SelectedText;
string fullBodyText = index.FullText;
```

Exclusions are opt-in. Region offsets refer to `FullText`, with reasons for plain-text quote prefixes, the `-- ` signature separator, HTML blockquotes, and supported mail-client quote/signature classes. Unclassified text does not establish authorship. The projection prefers the plain-text alternative by default, never opens attachments, rejects selected source bodies above `MaxSourceChars` (2 Mi characters by default), and clips projected text at `MaxTextChars` (256 Ki characters) without splitting a Unicode surrogate pair. `MaxProjectionChars` separately bounds generated HTML after text encoding or RTF conversion at 16 Mi characters. `Truncated` and `Diagnostics` report omitted text; the returned full text is still subject to that bound. A missing body produces empty indexing text and a diagnostic.

Use `EmailHtmlComposer` for replies and forwards with a rich quotation and plain-text alternative:

```csharp
EmailHtmlCompositionResult reply = EmailHtmlComposer.ReplyAll(
    message, new EmailAddress("me@example.test"), "Thanks for the update.");
reply.Document.Save("reply.eml");
```

The core composer owns recipients, own-address exclusions and threading. Authored text is HTML-encoded, and original markup passes through the shared untrusted policy. Automatic meta-refresh navigation is removed by the mail safety projection. Quoted resource references, images, media, form controls, styles and inline attachments are omitted; resource diagnostics describe the quotation policy. An oversized rich quotation becomes bounded text instead of partially clipped markup. `EmailHtmlCompositionOptions.Composition` controls quotation and threading, `MaxSourceChars` bounds the original body, and `MaxProjectionChars` separately bounds generated HTML. Missing bodies are not quoted. The result is an independent draft; saving it does not send mail.

For a plain-text share artifact, `EmailHtmlShareCopy` applies indexing before the core field-selection policy:

```csharp
EmailHtmlShareCopyResult share = EmailHtmlShareCopy.Create(message,
    new EmailShareCopyOptions { ReplacementSubject = "Shared discussion" },
    new EmailIndexTextOptions { ExcludeQuotes = true, ExcludeSignatures = true });
share.Copy.Document.Save("share.eml");
```

The result records the selected body representation, truncation, diagnostic codes and explicitly removed
concealment findings without returning original text or finding previews. An explicit `ReplacementBodyText`
bypasses body projection. Recognized quote/signature exclusions are heuristics and do not establish authorship.
Indexing does not remove every form of CSS-concealed text.

Inspect the original HTML with `OfficeIMO.Html.HtmlContentSafety.Inspect`, review its findings, and pass
an `OfficeIMO.ContentSafety.OfficeContentCleanupSelection` containing the chosen finding IDs as
`cleanupSelection` to remove those exact findings before projection. This uses the shared content-safety
owner and leaves the original body unchanged. Review retained text and selected attachment payloads before
sharing; this operation does not promise complete anonymization or retain original integrity assurance.

`EmailHtmlDataInspector.Inspect("message.eml")` combines the core mail-data report with inspection of the
original HTML through `HtmlContentSafety` and the shared email body's active-content policy. It reports concealed
text mechanisms and instruction-like risk classifications, plus counts of blocked elements and event-handler
attributes. Finding text, hashes and previews are omitted. The original body is left intact; no resource is fetched.

The HTML source is bounded to one million UTF-16 units by default. `HtmlInspectionStatus` distinguishes a completed
inspection from an oversized body, an unavailable inspection (including engine finding limits), a missing HTML
body and catalog-only inspection. Finding samples that reach their bound remain explicit. A completed HTML
inspection is not a malware verdict or proof that all content is safe; RTF/plain-text and unselected store items
are not qualified by this adapter. Cancellation is observed before and after the shared synchronous HTML passes.
