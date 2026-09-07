# A repeatable document workflow

A compact technical guide authored with the Markdown builder.

## 1. Describe the document

Keep data, document construction, and delivery as separate steps.

```csharp
var report = MarkdownDoc.Create()
    .H1("Weekly delivery")
    .P("Generated from application data.");
```

## 2. Check the output

| Check | Purpose |
| --- | --- |
| Source text | Keep an editable, portable input |
| PDF preview | Inspect page flow and typography |
| Output hash | Identify the generated artifact |

> [!TIP] Start small
> Build a representative document before adding a batch pipeline.

## 3. Share the result

- Keep the original Markdown alongside the PDF.
- Use the same guide in repository documentation and a report bundle.
