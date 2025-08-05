# OfficeIMO Converters - Master TODO

## ✅ Phase 1: Project Restructuring (COMPLETED)

### 1.1 Rename Existing Projects ✅
- [x] Renamed `OfficeIMO.Pdf` → `OfficeIMO.Word.Pdf`
- [x] Renamed `OfficeIMO.Markdown` → `OfficeIMO.Word.Markdown`
- [x] Renamed `OfficeIMO.Html` → `OfficeIMO.Word.Html`
- [x] Updated all project file names (.csproj)
- [x] Updated all namespaces to `OfficeIMO.Word.{Format}`
- [x] Updated solution file references

### 1.2 Package Configuration ✅
- [x] Added `Markdig` v0.41.3 to `OfficeIMO.Word.Markdown`
- [x] Added `AngleSharp` v1.3.0 to `OfficeIMO.Word.Html`  
- [x] Updated `QuestPDF` to v2025.7.0 in `OfficeIMO.Word.Pdf`
- [x] Updated PackageIds in all .csproj files
- [x] Updated descriptions in all .csproj files

### 1.3 Clean Architecture ✅
- [x] Removed IWordConverter interface usage
- [x] Removed ConverterRegistry usage
- [x] Created extension methods for each converter
- [x] Created Options classes for each converter

---

## 📋 Phase 2: Converter Implementation (IN PROGRESS)

### 2.1 Markdown Converter - `OfficeIMO.Word.Markdown`
- [x] Created `MarkdownToWordConverter.cs` with Markdig
- [x] Created `WordToMarkdownConverter.cs` (custom)
- [x] Created `MarkdownOptions.cs`
- [x] Created extension methods in `WordMarkdownConverterExtensions.cs`
- [ ] **Complete implementation of all Markdown features:**
  - [x] Headings (H1-H6)
  - [x] Paragraphs
  - [x] Bold/Italic/Emphasis
  - [x] Lists (ordered/unordered)
  - [x] Code blocks
  - [x] Inline code
  - [x] Tables (basic)
  - [ ] Links (full implementation with URL downloading)
  - [ ] Images (downloading and embedding)
  - [ ] Blockquotes (complete styling)
  - [ ] Horizontal rules
  - [ ] Checkbox lists
  - [ ] Nested lists (full support)
  - [ ] HTML in Markdown
  - [ ] Footnotes
  - [ ] Task lists

### 2.2 HTML Converter - `OfficeIMO.Word.Html`
- [x] Created `HtmlToWordConverter.cs` with AngleSharp
- [x] Created `WordToHtmlConverter.cs` with AngleSharp
- [x] Created `HtmlOptions.cs`
- [x] Created extension methods in `WordHtmlConverterExtensions.cs`
- [ ] **Complete implementation of all HTML features:**
  - [x] Headings (H1-H6)
  - [x] Paragraphs
  - [x] Basic text formatting (bold, italic, underline)
  - [x] Lists (ul/ol)
  - [x] Tables (basic)
  - [ ] Links (complete with all href types)
  - [ ] Images (downloading, base64, external)
  - [ ] Nested lists
  - [ ] Table colspan/rowspan
  - [ ] CSS styles parsing
  - [ ] Forms (read-only conversion)
  - [ ] SVG support
  - [ ] Iframe handling

### 2.3 PDF Converter - `OfficeIMO.Word.Pdf` ✅
- [x] Already has `WordPdfConverterExtensions.cs`
- [x] Uses QuestPDF
- [ ] **Enhancements needed:**
  - [ ] Clickable hyperlinks
  - [ ] Better image positioning
  - [ ] Table spanning support
  - [ ] Bookmarks/TOC navigation
  - [ ] Metadata support

---

## 🧪 Phase 3: Testing & Examples

### 3.1 Unit Tests
- [ ] Create `OfficeIMO.Tests/Word.Markdown.Tests.cs`
  - [ ] Test Markdown → Word conversion
  - [ ] Test Word → Markdown conversion
  - [ ] Test round-trip conversion
  - [ ] Test edge cases
  
- [ ] Create `OfficeIMO.Tests/Word.Html.Tests.cs`
  - [ ] Test HTML → Word conversion
  - [ ] Test Word → HTML conversion
  - [ ] Test round-trip conversion
  - [ ] Test malformed HTML handling

- [ ] Update `OfficeIMO.Tests/Word.Pdf.Tests.cs`
  - [ ] Test new features
  - [ ] Test complex documents

### 3.2 Examples
- [x] Created `ConversionExamples.cs` (basic structure)
- [x] Created `ReadmeConversionExample.cs` (analysis tool)
- [ ] Update examples to use new converters:
  - [ ] Remove all ConverterRegistry examples
  - [ ] Create simple Markdown examples
  - [ ] Create simple HTML examples
  - [ ] Create PDF examples
  - [ ] Create round-trip examples

### 3.3 Real-World Testing
- [ ] Test with README.md conversion
- [ ] Test with complex HTML documents
- [ ] Test with GitHub Flavored Markdown
- [ ] Test with various Word document structures

---

## 📚 Phase 4: Documentation

### 4.1 Package Documentation
- [ ] Update main README.md with new package structure
- [ ] Create README for `OfficeIMO.Word.Markdown`
- [ ] Create README for `OfficeIMO.Word.Html`
- [ ] Update README for `OfficeIMO.Word.Pdf`

### 4.2 Usage Documentation
- [ ] Document installation process
- [ ] Document basic usage examples
- [ ] Document options and configuration
- [ ] Document limitations and known issues

### 4.3 API Documentation
- [ ] Ensure XML documentation comments
- [ ] Generate API documentation
- [ ] Create migration guide from old structure

---

## 🚀 Phase 5: Release Preparation

### 5.1 Build & CI/CD
- [ ] Update build scripts for new structure
- [ ] Update CI/CD pipelines
- [ ] Test package generation
- [ ] Test cross-platform compatibility

### 5.2 NuGet Packages
- [ ] Finalize package metadata
- [ ] Test package installation
- [ ] Prepare release notes
- [ ] Version numbering strategy

### 5.3 Breaking Changes
- [ ] Document all breaking changes
- [ ] Provide migration path
- [ ] Update examples repository

---

## 🎯 Priority Items (Next Steps)

1. **High Priority - Core Functionality**
   - [ ] Complete image handling in Markdown converter
   - [ ] Complete link handling in HTML converter
   - [ ] Add table support with proper formatting
   - [ ] Fix round-trip conversion issues

2. **Medium Priority - Enhanced Features**
   - [ ] Add CSS parsing for HTML
   - [ ] Add nested list support
   - [ ] Add code syntax highlighting
   - [ ] Add image downloading

3. **Low Priority - Nice to Have**
   - [ ] SVG support
   - [ ] Footnotes
   - [ ] Custom styles preservation
   - [ ] Advanced table features

---

## 📝 Notes

- **Dependencies Updated**: All packages now use latest versions
- **Architecture**: Clean separation with extension methods
- **No Registry**: Removed complex ConverterRegistry pattern
- **Industry Standards**: Using Markdig for MD, AngleSharp for HTML, QuestPDF for PDF

---

## 🔧 Technical Debt

- [ ] Remove old converter implementations from `OfficeIMO.Word/Converters/`
- [ ] Clean up old examples using IWordConverter
- [ ] Remove duplicate converter files
- [ ] Consolidate helper methods

---

## Success Metrics

When complete, users should be able to:
```csharp
// Simple, clean API
using OfficeIMO.Word;
using OfficeIMO.Word.Markdown;

// Markdown operations
var doc = WordDocument.LoadFromMarkdown(File.ReadAllText("README.md"));
doc.SaveAsMarkdown("output.md");

// HTML operations  
using OfficeIMO.Word.Html;
doc.SaveAsHtml("output.html");

// PDF operations
using OfficeIMO.Word.Pdf;
doc.SaveAsPdf("output.pdf");
```

No complex patterns, just simple extension methods that work!