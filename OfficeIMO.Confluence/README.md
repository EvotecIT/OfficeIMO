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

Credentials remain caller-owned through `IConfluenceCredentialSource`. The library does not provide a secret store or persist tokens.
