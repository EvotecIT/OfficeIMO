# OfficeIMO Converters - Master TODO

## 🔴 CURRENT IMPLEMENTATION STATUS

### What Actually Works:
- ✅ Project builds and compiles
- ✅ Extension methods work (`string.LoadFromHtml()`, `string.LoadFromMarkdown()`)
- ✅ Basic project structure with separate packages
- ✅ Null checking and exception handling
- ⚠️ **PDF converter** - Partially working (see details below)

### What Does NOT Work (Despite Being Marked as Done):
- ❌ **HTML to Word** - Only extracts plain text, no formatting
- ❌ **Word to HTML** - Only creates basic `<p>` tags, no formatting
- ❌ **Markdown to Word** - Not using Markdig at all, just splits text by lines
- ❌ **Word to Markdown** - Only extracts plain text, no markdown formatting
- ❌ **ALL formatting features** - Bold, italic, headings, lists, tables, links - NONE work

### PDF Converter Status (Partially Working):
#### What PDF Does Right:
- ✅ Uses `document.Elements` to iterate through content
- ✅ Handles WordParagraph and WordTable with pattern matching
- ✅ Pre-processes lists via `document.Lists` to add markers
- ✅ Renders default headers/footers
- ✅ Handles nested tables

#### What PDF Does Wrong:
- ❌ Doesn't iterate through sections (uses aggregated `document.Elements`)
- ❌ Only handles Default headers/footers (ignores First/Even)
- ❌ Missing many element types (WordImage, WordHyperLink, etc.)

## 📁 Project Structure
```
OfficeIMO/
├── OfficeIMO.Word/              # Core Word library
├── OfficeIMO.Word.Pdf/          # PDF converter (WORKING)
├── OfficeIMO.Word.Markdown/     # Markdown converter (STUB ONLY)
├── OfficeIMO.Word.Html/         # HTML converter (STUB ONLY)
├── OfficeIMO.Examples/          # Examples
└── OfficeIMO.Tests/             # Tests (mostly skipped)
```

## 🚨 PHASE 1: Fix What's Broken

### HTML to Word Converter (`HtmlToWordConverter.cs`)

#### Currently Implemented (Partially):
- ✅ AngleSharp HTML parsing setup
- ✅ Extracts text from `<p>` tags (text only, no formatting)
- ✅ Extracts text from `<h1>-<h6>` tags (text only, NO styling applied)
- ✅ Page settings from options

#### NOT Implemented (Need to do):
- ❌ Apply heading styles (currently line 65 has TODO comment)
- ❌ Bold/italic/underline formatting (`<b>`, `<i>`, `<u>`, `<strong>`, `<em>`)
- ❌ Hyperlinks (`<a href="">`)
- ❌ Lists (`<ul>`, `<ol>`, `<li>`)
- ❌ Tables (`<table>`, `<tr>`, `<td>`)
- ❌ Images (`<img>`)
- ❌ Line breaks (`<br>`)
- ❌ CSS styles (inline or classes)
- ❌ Nested structures (nested lists, nested tables)

### Word to HTML Converter (`WordToHtmlConverter.cs`)

#### Currently Implemented:
- ✅ Basic HTML document structure
- ✅ UTF-8 charset meta tag
- ✅ Extracts paragraphs as simple `<p>` elements (text only)

#### NOT Implemented:
- ❌ Heading detection and conversion to `<h1>-<h6>`
- ❌ Bold/italic/underline formatting
- ❌ Hyperlinks
- ❌ Lists (bullet and numbered)
- ❌ Tables
- ❌ Images
- ❌ CSS styles
- ❌ Document metadata (title, etc.)

### Markdown to Word Converter (`MarkdownToWordConverter.cs`)

#### Currently Implemented:
- ✅ Splits markdown by newlines
- ✅ Adds each line as a paragraph
- ✅ Page settings from options

#### NOT Implemented:
- ❌ **NOT USING MARKDIG AT ALL!** (despite having the dependency)
- ❌ Heading parsing (`#`, `##`, etc.)
- ❌ Bold parsing (`**text**`)
- ❌ Italic parsing (`*text*`)
- ❌ Link parsing (`[text](url)`)
- ❌ List parsing (`-`, `*`, `1.`)
- ❌ Code block parsing (` ``` `)
- ❌ Inline code parsing (`` ` ``)
- ❌ Table parsing (pipe syntax)
- ❌ Image parsing (`![alt](url)`)
- ❌ Blockquote parsing (`>`)

### Word to Markdown Converter (`WordToMarkdownConverter.cs`)

#### Currently Implemented:
- ✅ Extracts paragraph text

#### NOT Implemented:
- ❌ Heading detection and conversion to `#` syntax
- ❌ Bold detection and conversion to `**text**`
- ❌ Italic detection and conversion to `*text*`
- ❌ Hyperlink conversion to `[text](url)`
- ❌ List conversion (bullet and numbered)
- ❌ Table conversion to pipe syntax
- ❌ Image conversion to `![alt](url)`

## 🔧 PHASE 2: Implementation Priority

### Week 1: Get Basic Features Working

#### Day 1-2: Fix Markdown Converter (USE MARKDIG!)
```csharp
// MarkdownToWordConverter.cs - THIS IS WHAT NEEDS TO BE DONE:
var pipeline = new MarkdownPipelineBuilder()
    .UseAdvancedExtensions()
    .Build();
var markdownDocument = Markdig.Markdown.Parse(markdown, pipeline);

// Walk the AST and convert to Word elements
foreach (var block in markdownDocument) {
    switch (block) {
        case HeadingBlock heading:
            // Apply actual heading style!
            var para = wordDoc.AddParagraph(heading.Inline.FirstChild.ToString());
            para.Style = GetWordHeadingStyle(heading.Level);
            break;
        case ListBlock list:
            // Create actual Word list!
            var wordList = wordDoc.AddList();
            // ... add items
            break;
        // etc...
    }
}
```

#### Day 3-4: Fix HTML Converter (USE ANGLESHARP PROPERLY!)
```csharp
// HtmlToWordConverter.cs - THIS IS WHAT NEEDS TO BE DONE:
foreach (var element in document.Body.Children) {
    switch (element.TagName.ToLower()) {
        case "h1":
        case "h2":
        // ... etc
            var level = int.Parse(element.TagName.Substring(1));
            var para = wordDoc.AddParagraph(element.TextContent);
            para.Style = GetWordHeadingStyle(level); // ACTUALLY SET THE STYLE!
            break;
        case "p":
            // Handle formatting within paragraph
            var para = wordDoc.AddParagraph();
            ProcessInlineElements(element, para); // Handle <b>, <i>, etc.
            break;
        case "ul":
        case "ol":
            // Create actual list!
            var list = wordDoc.AddList();
            ProcessListItems(element, list);
            break;
    }
}
```

#### Day 5: Fix Tests
- Remove Skip attributes as features are implemented
- Ensure tests actually pass

### Week 2: Add Missing Features

#### Essential Features to Add:
1. **Links/Hyperlinks** - Both HTML and Markdown
2. **Bold/Italic formatting** - Both directions
3. **Lists** - Proper nested list support
4. **Tables** - Basic table structure
5. **Images** - At least base64 support

## 📝 Examples That Need to be Created/Fixed

### Current State of Examples:
- Many examples were deleted during restructuring
- Need to restore/recreate basic examples

### Examples to Create:
```csharp
// 1. Basic HTML Example
public static void Example_HtmlBasics() {
    // This should actually work with formatting!
    string html = "<h1>Title</h1><p>This is <b>bold</b> and <i>italic</i></p>";
    var doc = html.LoadFromHtml();
    doc.Save("output.docx");
    
    // Verify the heading is actually styled
    Assert.That(doc.Paragraphs[0].Style == WordParagraphStyles.Heading1);
}

// 2. Basic Markdown Example  
public static void Example_MarkdownBasics() {
    // This should use Markdig and work!
    string markdown = "# Heading\n\nThis is **bold** and *italic*";
    var doc = markdown.LoadFromMarkdown();
    doc.Save("output.docx");
    
    // Verify formatting is applied
    Assert.That(doc.Paragraphs[0].Style == WordParagraphStyles.Heading1);
}
```

## 🧪 Tests Currently Skipped (Need Implementation)

### HTML Tests (All Skipped):
- `Test_Html_RoundTrip` - Needs formatting implementation
- `Test_Html_Headings_RoundTrip` - Needs heading styles
- `Test_Html_Lists_RoundTrip` - Needs list implementation
- `Test_Html_Table_RoundTrip` - Needs table implementation
- `Test_Html_NestedTable_RoundTrip` - Needs nested table support
- `Test_Html_Image_Base64_RoundTrip` - Needs image support
- `Test_Html_Image_File_RoundTrip` - Needs image support
- `Test_Html_FontResolver` - Needs font mapping
- `Test_Html_Urls_CreateHyperlinks` - Needs hyperlink support
- `Test_Html_InlineStyles_ParagraphStyle` - Needs CSS parsing

### Markdown Tests (All Skipped):
- `Test_Markdown_RoundTrip` - Needs Markdig implementation
- `Test_Markdown_Lists_RoundTrip` - Needs list parsing
- `Test_Markdown_FontResolver` - Needs font mapping
- `Test_Markdown_Urls_CreateHyperlinks` - Needs link parsing

## 🎯 Definition of "DONE"

A feature is ONLY considered done when:
1. ✅ It actually converts the format correctly (not just extracts text)
2. ✅ Formatting is preserved (bold, italic, etc.)
3. ✅ Structure is preserved (headings have styles, lists are lists, etc.)
4. ✅ Tests pass without Skip attribute
5. ✅ Round-trip works (Format → Word → Format preserves content)
6. ✅ Example demonstrating the feature exists and works

## ⚠️ Critical Issues to Fix

1. **Markdig is not being used at all** - The Markdown converter just splits by newlines
2. **Heading styles are not applied** - Headings are extracted but not styled
3. **No formatting is preserved** - Bold, italic, etc. are completely ignored
4. **Lists don't work** - No list parsing or creation
5. **Links don't work** - No hyperlink support
6. **Images don't work** - No image handling

## 📋 Understanding OfficeIMO.Word Structure

### Document Hierarchy:
```csharp
WordDocument
├── Sections[] (document can have multiple sections)
│   ├── Elements[] (List<WordElement> - all content in flow)
│   │   ├── WordParagraph : WordElement
│   │   ├── WordTable : WordElement
│   │   ├── WordList : WordElement
│   │   ├── WordImage : WordElement
│   │   ├── WordHyperLink : WordElement
│   │   └── ... (many other types inherit from WordElement)
│   ├── Header (WordHeaders)
│   │   ├── Default (WordHeader : WordHeaderFooter)
│   │   │   ├── Paragraphs[]
│   │   │   ├── Tables[]
│   │   │   └── Lists[]
│   │   ├── First (WordHeader)
│   │   └── Even (WordHeader)
│   └── Footer (WordFooters)
│       ├── Default (WordFooter : WordHeaderFooter)
│       ├── First (WordFooter)
│       └── Even (WordFooter)
└── Properties (metadata, settings, etc.)
```

### Correct Way to Process Document Content:
```csharp
// Process main document body
foreach (var section in document.Sections) {
    foreach (var element in section.Elements) {
        switch (element) {
            case WordParagraph para:
                // Check para.Style for headings
                // Check para.Bold, para.Italic for formatting
                // Check para.Text for content
                break;
            case WordTable table:
                // Process table rows and cells
                break;
            case WordList list:
                // Process list items
                break;
            case WordHyperLink link:
                // Process hyperlinks
                break;
            // ... handle other types
        }
    }
    
    // Process headers if they exist
    if (section.Header?.Default != null) {
        // Process section.Header.Default.Paragraphs, Tables, etc.
    }
    
    // Process footers if they exist  
    if (section.Footer?.Default != null) {
        // Process section.Footer.Default.Paragraphs, Tables, etc.
    }
}
```

## 📋 Helper Methods Needed in OfficeIMO.Word

Before implementing converters, we may need to add these to the main API:

```csharp
// Style helpers
WordParagraphStyles GetHeadingStyleForLevel(int level);
int GetLevelForHeadingStyle(WordParagraphStyles style);

// Formatting helpers  
WordParagraph AddFormattedText(string text, bool bold, bool italic);
WordHyperlink AddHyperlink(string text, string url); // May already exist

// List helpers
WordList CreateBulletList(); // Check if AddList() already exists
WordList CreateNumberedList();

// Image helpers
WordImage AddImageFromBase64(string base64);
WordImage AddImageFromUrl(string url);
```

## 📝 Code Style Preferences

### Use Partial Classes for Large Files
**IMPORTANT:** Break large converter classes into logical partial classes instead of having one massive file:

```csharp
// Instead of one huge HtmlToWordConverter.cs with 1000+ lines, use:
HtmlToWordConverter.cs           // Main class definition and core logic
HtmlToWordConverter.Tables.cs    // Table handling methods
HtmlToWordConverter.Lists.cs     // List handling methods  
HtmlToWordConverter.Images.cs    // Image handling methods
HtmlToWordConverter.Styles.cs    // CSS and styling methods
```

This pattern is already used throughout OfficeIMO.Word:
- `WordDocument.cs`, `WordDocument.Images.cs`, `WordDocument.Tables.cs`, etc.
- `WordSection.cs`, `WordSection.PublicMethods.cs`, `WordSection.PrivateMethods.cs`
- `WordPdfConverterExtensions.cs`, `WordPdfConverterExtensions.Rendering.cs`, `WordPdfConverterExtensions.Helpers.cs`

### Benefits:
- Easier to navigate and maintain
- Better organization of related functionality
- Reduces merge conflicts in team development
- Follows existing OfficeIMO patterns

## 🚀 Next Immediate Steps

1. **STOP claiming features are implemented when they're not**
2. **START with Markdown converter** - Actually use Markdig
3. **FIX heading styles** - They should actually apply WordParagraphStyles
4. **IMPLEMENT basic formatting** - Bold and italic at minimum
5. **CREATE working examples** - Not just stub code
6. **ENABLE tests** - As features are actually implemented
7. **USE PARTIAL CLASSES** - Break converters into logical parts as they grow

## 📊 Real Progress Tracking

### Markdown Converter:
- [ ] Parse with Markdig (NOT DONE)
- [ ] Headings with styles (NOT DONE)
- [ ] Bold/Italic (NOT DONE)
- [ ] Lists (NOT DONE)
- [ ] Links (NOT DONE)
- [ ] Code blocks (NOT DONE)
- [ ] Tables (NOT DONE)
- [ ] Images (NOT DONE)

### HTML Converter:
- [ ] Headings with styles (NOT DONE - only text extraction)
- [ ] Bold/Italic/Underline (NOT DONE)
- [ ] Hyperlinks (NOT DONE)
- [ ] Lists (NOT DONE)
- [ ] Tables (NOT DONE)
- [ ] Images (NOT DONE)
- [ ] CSS styles (NOT DONE)

### Word to HTML:
- [ ] Detect heading styles (NOT DONE)
- [ ] Export formatting (NOT DONE)
- [ ] Export hyperlinks (NOT DONE)
- [ ] Export lists (NOT DONE)
- [ ] Export tables (NOT DONE)
- [ ] Export images (NOT DONE)

### Word to Markdown:
- [ ] Detect heading styles (NOT DONE)
- [ ] Export formatting (NOT DONE)
- [ ] Export hyperlinks (NOT DONE)
- [ ] Export lists (NOT DONE)
- [ ] Export tables (NOT DONE)

## 🏁 Success Criteria

The converters are ONLY complete when this code actually works:

```csharp
// This should preserve formatting, not just extract text!
string markdown = "# Heading\n\n**Bold** and *italic* and [link](http://example.com)";
var doc = markdown.LoadFromMarkdown();

// These assertions should pass:
Assert.That(doc.Paragraphs[0].Style == WordParagraphStyles.Heading1);
Assert.That(doc.Paragraphs[1].Text.Contains("Bold"));
// Bold should actually be bold, italic should be italic, link should be a hyperlink

// Round trip should work:
string markdownOut = doc.ToMarkdown();
Assert.That(markdownOut.Contains("# Heading"));
Assert.That(markdownOut.Contains("**Bold**"));
Assert.That(markdownOut.Contains("*italic*"));
Assert.That(markdownOut.Contains("[link](http://example.com)"));
```

Until the above code works, the converters are NOT implemented, just stubbed.