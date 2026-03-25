---
title: Markdown Builder
description: Fluent API for building Markdown documents -- headings, paragraphs, lists, tables, code blocks, images, and callouts.
order: 41
---

# Markdown Builder

The `MarkdownDoc` class provides a fluent API for composing Markdown documents entirely in code. Every method returns the same `MarkdownDoc` instance for chaining. The resulting document renders to GitHub-flavored Markdown via `ToString()`.

## Headings

```csharp
using OfficeIMO.Markdown;

var doc = MarkdownDoc.Create()
    .H1("Main Title")
    .H2("Section")
    .H3("Subsection")
    .H4("Sub-subsection")
    .H5("Minor heading")
    .H6("Smallest heading");

// Output:
// # Main Title
// ## Section
// ### Subsection
// ...
```

## Paragraphs

### Simple Text

```csharp
doc.P("This is a simple paragraph of text.");
```

### Formatted Paragraphs

Use the `ParagraphBuilder` for inline formatting:

```csharp
doc.P(p => p
    .Text("This is ")
    .Bold("important")
    .Text(" and this is ")
    .Italic("emphasized")
    .Text(". You can also use ")
    .Code("inline code")
    .Text(" and ")
    .Link("links", "https://example.com")
    .Text(".")
);
```

## Lists

### Unordered Lists

```csharp
doc.Ul(ul => ul
    .Item("First item")
    .Item("Second item")
    .Item("Third item with nested list", nested => nested
        .Item("Nested item A")
        .Item("Nested item B")
    )
);
```

Output:

```markdown
- First item
- Second item
- Third item with nested list
    - Nested item A
    - Nested item B
```

### Ordered Lists

```csharp
doc.Ol(ol => ol
    .Item("Step one")
    .Item("Step two")
    .Item("Step three")
);
```

Output:

```markdown
1. Step one
2. Step two
3. Step three
```

### Lists from Collections

```csharp
var features = new[] { "Cross-platform", "No dependencies", "Fluent API" };
doc.Ul(features);
```

### Definition Lists

```csharp
doc.Dl(dl => dl
    .Term("API", "Application Programming Interface")
    .Term("SDK", "Software Development Kit")
);
```

## Tables

### Manual Table Construction

```csharp
doc.Table(t => t
    .Header("Name", "Role", "Location")
    .Row("Alice", "Developer", "New York")
    .Row("Bob", "Designer", "London")
    .Row("Carol", "Manager", "Tokyo")
);
```

Output:

```markdown
| Name  | Role      | Location |
|-------|-----------|----------|
| Alice | Developer | New York |
| Bob   | Designer  | London   |
| Carol | Manager   | Tokyo    |
```

### Table from Objects

Generate a table from any collection of objects:

```csharp
var data = new[] {
    new { Product = "Widget A", Price = 9.99, Stock = 150 },
    new { Product = "Widget B", Price = 14.99, Stock = 75 },
};

doc.TableFrom(data);
```

### Table with Auto-Alignment

Automatically align numeric columns to the right and date columns to the center:

```csharp
doc.TableFromAuto(data, alignNumeric: true, alignDates: true);
```

### Table with Column Selectors

```csharp
doc.TableFrom(data,
    ("Product", x => x.Product),
    ("Price ($)", x => x.Price.ToString("F2")),
    ("In Stock", x => x.Stock > 100 ? "Yes" : "Low")
);
```

### Column Alignment

```csharp
doc.Table(t => t
    .Header("Left", "Center", "Right")
    .Align(ColumnAlignment.Left, ColumnAlignment.Center, ColumnAlignment.Right)
    .Row("A", "B", "C")
);
```

## Code Blocks

### Fenced Code Blocks

```csharp
doc.Code("csharp", @"var greeting = ""Hello, World!"";
Console.WriteLine(greeting);");
```

Output:

````markdown
```csharp
var greeting = "Hello, World!";
Console.WriteLine(greeting);
```
````

### Code with Captions

```csharp
doc.Code("json", @"{""name"": ""OfficeIMO"", ""version"": ""1.0.38""}")
   .Caption("package.json configuration");
```

## Block Quotes

```csharp
// Simple quote
doc.Quote("The best way to predict the future is to invent it.");

// Multi-line quote with builder
doc.Quote(q => q
    .Line("First line of the quote.")
    .Line("Second line of the quote.")
    .Attribution("Alan Kay")
);
```

## Callouts / Admonitions

```csharp
doc.Callout("note", "Important", "This feature requires .NET 8 or later.");
doc.Callout("warning", "Breaking Change", "The API signature changed in v2.0.");
doc.Callout("tip", "Performance", "Use streaming mode for files over 100MB.");
```

Output:

```markdown
> [!NOTE]
> **Important**
> This feature requires .NET 8 or later.
```

## Images

```csharp
doc.Image("screenshot.png", alt: "Dashboard screenshot", title: "Dashboard");

// With explicit dimensions
doc.Image("logo.svg", alt: "Logo", width: 200, height: 60);
```

## Horizontal Rules

```csharp
doc.Hr();
```

## Collapsible Details

```csharp
doc.Details("Click to expand", body => body
    .P("This content is hidden by default.")
    .Code("bash", "echo 'Revealed!'")
);
```

Output:

```markdown
<details>
<summary>Click to expand</summary>

This content is hidden by default.

```bash
echo 'Revealed!'
```

</details>
```

## Front Matter

Add YAML front matter to the document:

```csharp
doc.FrontMatter(new {
    title = "My Document",
    date = "2025-01-15",
    tags = new[] { "docs", "guide" }
});
```

Output:

```markdown
---
title: My Document
date: 2025-01-15
tags:
  - docs
  - guide
---
```

## Rendering to String

```csharp
string markdown = doc.ToString();
File.WriteAllText("output.md", markdown);
```

## Rendering to HTML

```csharp
string html = doc.ToHtml(new HtmlOptions {
    Style = HtmlStyle.GitHub,
    IncludeTableOfContents = true,
    PrismOptions = new PrismOptions { Enabled = true }
});
```
