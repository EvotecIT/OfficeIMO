# OfficeIMO.Markdown.Html

HTML to Markdown conversion for `OfficeIMO.Markdown`.

## Usage

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

var markdown = "<h1>Hello</h1><p>Body</p>".ToMarkdown();
var document = "<h1>Hello</h1><p>Body</p>".LoadFromHtml();
```
