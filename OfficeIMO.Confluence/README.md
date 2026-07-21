# OfficeIMO.Confluence

`OfficeIMO.Confluence` is a dependency-light Confluence Cloud client with ADF, Markdown, HTML, attachment, dry-run, and managed-section support.

```csharp
var session = new ConfluenceSession(
    new ConfluenceBasicCredentialSource(email, apiToken),
    new ConfluenceSessionOptions { SiteUri = new Uri("https://example.atlassian.net/") });

using ConfluenceClient client = session.CreateClient();
ConfluencePage page = await client.GetPageAsync("123", ConfluenceBodyFormat.AtlasDocFormat);
ConfluenceContentConversionResult<string> markdown = ConfluenceContentConverter.ToMarkdown(page);
```

Use `ConfluenceClient.PlanCreatePage` and `PlanUpdatePage` for non-executing write previews. `ConfluenceManagedSection.Apply` updates only a stable marker pair and returns before/after SHA-256 hashes. Safe reads retry transient failures; creates, updates, and attachment uploads are not automatically retried after an ambiguous write.

OAuth 2.0 bearer sessions set `ConfluenceSessionOptions.CloudId`; the client then retains the Atlassian gateway prefix for every request:

```csharp
var oauthSession = new ConfluenceSession(
    new ConfluenceBearerCredentialSource(accessToken),
    new ConfluenceSessionOptions {
        SiteUri = new Uri("https://example.atlassian.net/"),
        CloudId = cloudId,
    });
```

Use `PlanDeletePage` before `DeletePageAsync` when a caller needs a reviewable destructive-operation plan. Attachment overloads accepting `Stream` avoid buffering file-sized payloads; caller-owned streams remain open.

Credentials remain caller-owned through `IConfluenceCredentialSource`. The library does not provide a secret store or persist tokens.
