using OfficeIMO.Markdown;
using OfficeIMO.OneNote.Html;
using OfficeIMO.OneNote.Markdown;
using OfficeIMO.OneNote.Pdf;
using System.Text;

namespace OfficeIMO.OneNote.Tests;

public sealed class ConverterTests {
    [Fact]
    public void SharedProjectionCoversHierarchyContentAndOptionalRelatedPages() {
        OneNoteNotebook notebook = CreateNotebook();
        int assetIndex = 0;

        string currentOnly = notebook.ToMarkdown(new OneNoteMarkdownOptions {
            AssetUriResolver = _ => "assets/item-" + (++assetIndex) + ".bin"
        });

        Assert.Contains("# Offline notebook", currentOnly);
        Assert.Contains("## Group A", currentOnly);
        Assert.Contains("### Section A", currentOnly);
        Assert.Contains("#### Current page", currentOnly);
        Assert.Contains("**Bold**", currentOnly);
        Assert.Contains("[link](https://example.com/a%20b)", currentOnly);
        Assert.Contains("- Item", currentOnly);
        Assert.Contains("| Column 1 | Column 2 |", currentOnly);
        Assert.Contains("![Diagram](assets/item-1.bin)", currentOnly);
        Assert.Contains("[sample.zip](assets/item-2.bin)", currentOnly);
        Assert.Contains("```math", currentOnly);
        Assert.Contains("x^2", currentOnly);
        Assert.DoesNotContain("Conflict: Conflict copy", currentOnly);
        Assert.DoesNotContain("Version: Historical copy", currentOnly);
        Assert.Equal(2, assetIndex);

        string withRelated = notebook.ToMarkdown(new OneNoteMarkdownOptions {
            IncludeConflictPages = true,
            IncludeVersionHistory = true
        });

        Assert.Contains("Conflict: Conflict copy", withRelated);
        Assert.Contains("Version: Historical copy", withRelated);
        Assert.Contains("Version: Conflict history", withRelated);
        Assert.Contains("Conflict: Historical conflict", withRelated);
    }

    [Fact]
    public void ProjectionProducesReusableMarkdownDocumentAndUtf8Bytes() {
        OneNoteSection section = CreateNotebook().SectionGroups[0].Sections[0];

        MarkdownDoc document = section.ToMarkdownDocument();
        string markdown = section.ToMarkdown();
        byte[] bytes = section.ToMarkdownBytes();

        Assert.Contains("Section A", markdown);
        Assert.Contains("Current page", document.ToMarkdown());
        Assert.Equal(markdown, new UTF8Encoding(false).GetString(bytes));
        Assert.Throws<ArgumentOutOfRangeException>(() => section.ToMarkdown(new OneNoteMarkdownOptions { HeadingLevel = 0 }));
    }

    [Fact]
    public void NotebookProjectionHonorsInterleavedTableOfContentsOrderAtEveryLevel() {
        var notebook = new OneNoteNotebook { Name = "Ordered" };
        OneNoteSection middle = SectionWithPage("Middle", "Middle page");
        middle.TableOfContentsOrder = 1;
        notebook.Sections.Add(middle);

        var last = new OneNoteSectionGroup { Name = "Last", TableOfContentsOrder = 2 };
        last.Sections.Add(SectionWithPage("Last section", "Last page"));
        notebook.SectionGroups.Add(last);

        var first = new OneNoteSectionGroup { Name = "First", TableOfContentsOrder = 0 };
        OneNoteSection nestedMiddle = SectionWithPage("Nested middle", "Nested middle page");
        nestedMiddle.TableOfContentsOrder = 1;
        first.Sections.Add(nestedMiddle);
        var nestedFirst = new OneNoteSectionGroup { Name = "Nested first", TableOfContentsOrder = 0 };
        nestedFirst.Sections.Add(SectionWithPage("Nested first section", "Nested first page"));
        first.SectionGroups.Add(nestedFirst);
        notebook.SectionGroups.Add(first);

        string markdown = notebook.ToMarkdown();

        Assert.True(markdown.IndexOf("## First", StringComparison.Ordinal) < markdown.IndexOf("## Middle", StringComparison.Ordinal));
        Assert.True(markdown.IndexOf("## Middle", StringComparison.Ordinal) < markdown.IndexOf("## Last", StringComparison.Ordinal));
        Assert.True(markdown.IndexOf("### Nested first", StringComparison.Ordinal) < markdown.IndexOf("### Nested middle", StringComparison.Ordinal));
    }

    [Fact]
    public async Task HtmlConversionSupportsDocumentFragmentBytesAndCallerOwnedStreams() {
        OneNoteSection section = CreateNotebook().SectionGroups[0].Sections[0];

        string document = section.ToHtmlDocument(htmlOptions: new HtmlOptions { AssetMode = AssetMode.Offline });
        string fragment = section.ToHtmlFragment(htmlOptions: new HtmlOptions { AssetMode = AssetMode.Offline });
        byte[] bytes = section.ToHtmlBytes(htmlOptions: new HtmlOptions { AssetMode = AssetMode.Offline });
        using var stream = new MemoryStream();
        await section.SaveAsHtmlAsync(stream, htmlOptions: new HtmlOptions { AssetMode = AssetMode.Offline });

        Assert.Contains("<!DOCTYPE html>", document, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Current page", document);
        Assert.Contains("<strong>Bold</strong>", fragment);
        Assert.DoesNotContain("<!DOCTYPE html>", fragment, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("Current page", new UTF8Encoding(false).GetString(bytes));
        Assert.True(stream.CanWrite);
        Assert.Contains("Current page", new UTF8Encoding(false).GetString(stream.ToArray()));
    }

    [Fact]
    public async Task PdfConversionProducesValidBytesAndLeavesCallerOwnedStreamOpen() {
        OneNoteSection section = CreateNotebook().SectionGroups[0].Sections[0];

        byte[] bytes = section.ToPdf();
        using var stream = new MemoryStream();
        await section.SaveAsPdfAsync(stream);

        Assert.True(bytes.Length > 100);
        Assert.Equal("%PDF", Encoding.ASCII.GetString(bytes, 0, 4));
        Assert.True(stream.CanWrite);
        Assert.Equal("%PDF", Encoding.ASCII.GetString(stream.ToArray(), 0, 4));
    }

    [Fact]
    public void PlainTextProjectionIsSharedForPagesElementsAndCells() {
        OneNoteSection section = CreateNotebook().SectionGroups[0].Sections[0];
        OneNotePage page = section.Pages[0];
        OneNoteTable table = Assert.IsType<OneNoteTable>(page.DirectContent[1]);

        Assert.Contains("Current page", OneNoteMarkdownProjection.ToText(page));
        Assert.Equal("Left", OneNoteMarkdownProjection.ToText(table.Rows[0].Cells[0]));
        Assert.Equal("Bold link", OneNoteMarkdownProjection.ToText(page.DirectContent[0]));
    }

    [Fact]
    public void ProjectionNormalizesRichEditControlsAndUnicodeNoncharactersWithoutMutatingSource() {
        const string nativeText = "Alpha\vBeta\u0001\uFDDF";
        var section = new OneNoteSection { Name = "Controls" };
        var page = new OneNotePage { Title = "Native\vtitle" };
        OneNoteParagraph paragraph = Paragraph(nativeText);
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        string text = OneNoteMarkdownProjection.ToText(page);
        string markdown = section.ToMarkdown();
        string html = section.ToHtmlDocument(htmlOptions: new HtmlOptions { AssetMode = AssetMode.Offline });
        byte[] pdf = section.ToPdf();

        Assert.Contains("Native\ntitle", text, StringComparison.Ordinal);
        Assert.Contains("Alpha\nBeta??", text, StringComparison.Ordinal);
        Assert.Contains("Native<br>title", markdown, StringComparison.Ordinal);
        Assert.Contains("Alpha<br>Beta??", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain('\v', html);
        Assert.DoesNotContain('\u0001', html);
        Assert.DoesNotContain('\uFDDF', html);
        Assert.Equal("%PDF", Encoding.ASCII.GetString(pdf, 0, 4));
        Assert.Equal(nativeText, Assert.Single(paragraph.Runs).Text);
        Assert.Equal("Native\vtitle", page.Title);
    }

    [Fact]
    public void SharedProjectionBoundsCallerSuppliedListIndentationAcrossConverters() {
        var section = new OneNoteSection { Name = "Lists" };
        var page = new OneNotePage { Title = "Extreme list" };
        var paragraph = new OneNoteParagraph {
            List = new OneNoteListInfo { Level = int.MaxValue }
        };
        paragraph.Runs.Add(new OneNoteTextRun { Text = "Bounded item" });
        page.DirectContent.Add(paragraph);
        section.Pages.Add(page);

        string markdown = section.ToMarkdown();
        string html = section.ToHtmlDocument(htmlOptions: new HtmlOptions { AssetMode = AssetMode.Offline });
        byte[] pdf = section.ToPdf();
        string expectedPrefix = new string(' ', OneNoteListInfo.MaxLevel * 2) + "- Bounded item";

        Assert.Contains(expectedPrefix, markdown, StringComparison.Ordinal);
        Assert.Contains("Bounded item", html, StringComparison.Ordinal);
        Assert.Equal("%PDF", Encoding.ASCII.GetString(pdf, 0, 4));
    }

    [Fact]
    public void RecursiveContentIsRejectedAcrossTextMarkdownHtmlAndPdfProjection() {
        var section = new OneNoteSection { Name = "Cycle" };
        var page = new OneNotePage { Title = "Recursive" };
        var outline = new OneNoteOutline();
        outline.Children.Add(outline);
        page.Outlines.Add(outline);
        section.Pages.Add(page);

        AssertProjectionError("ONENOTE_PROJECTION_CONTENT_CYCLE", () => OneNoteMarkdownProjection.ToText(outline));
        AssertProjectionError("ONENOTE_PROJECTION_CONTENT_CYCLE", () => OneNoteMarkdownProjection.ToMarkdown(outline));
        AssertProjectionError("ONENOTE_PROJECTION_CONTENT_CYCLE", () => section.ToMarkdown());
        AssertProjectionError("ONENOTE_PROJECTION_CONTENT_CYCLE", () => section.ToHtmlDocument());
        AssertProjectionError("ONENOTE_PROJECTION_CONTENT_CYCLE", () => section.ToPdf());
    }

    [Fact]
    public void NotebookProjectionRejectsCyclicAndExcessivelyDeepSectionGroups() {
        var cyclicNotebook = new OneNoteNotebook { Name = "Cyclic" };
        var cyclicGroup = new OneNoteSectionGroup { Name = "Loop" };
        cyclicGroup.SectionGroups.Add(cyclicGroup);
        cyclicNotebook.SectionGroups.Add(cyclicGroup);

        AssertProjectionError("ONENOTE_PROJECTION_GROUP_CYCLE", () => cyclicNotebook.ToMarkdown());

        var deepNotebook = new OneNoteNotebook { Name = "Deep" };
        var parent = new OneNoteSectionGroup { Name = "Parent" };
        parent.SectionGroups.Add(new OneNoteSectionGroup { Name = "Child" });
        deepNotebook.SectionGroups.Add(parent);

        AssertProjectionError(
            "ONENOTE_PROJECTION_GROUP_DEPTH",
            () => deepNotebook.ToMarkdown(new OneNoteMarkdownOptions { MaxSectionGroupDepth = 1 }));
    }

    [Fact]
    public void RelatedPageCyclesRemainIgnoredUntilTheirProjectionIsRequested() {
        var section = new OneNoteSection { Name = "Related" };
        var conflict = new OneNotePage { Title = "Conflict root" };
        conflict.ConflictPages.Add(conflict);
        section.Pages.Add(conflict);
        var version = new OneNotePage { Title = "Version root" };
        version.VersionHistory.Add(version);
        section.Pages.Add(version);

        string currentOnly = section.ToMarkdown();

        Assert.Contains("Conflict root", currentOnly, StringComparison.Ordinal);
        Assert.Contains("Version root", currentOnly, StringComparison.Ordinal);
        AssertProjectionError(
            "ONENOTE_PROJECTION_PAGE_CYCLE",
            () => section.ToMarkdown(new OneNoteMarkdownOptions { IncludeConflictPages = true }));
        AssertProjectionError(
            "ONENOTE_PROJECTION_PAGE_CYCLE",
            () => section.ToMarkdown(new OneNoteMarkdownOptions { IncludeVersionHistory = true }));
    }

    [Fact]
    public void ProjectionRejectsConfiguredPageAndContentDepthOverruns() {
        var page = new OneNotePage { Title = "Deep page" };
        page.ConflictPages.Add(new OneNotePage { Title = "Nested" });
        var section = new OneNoteSection { Name = "Depth" };
        section.Pages.Add(page);
        AssertProjectionError(
            "ONENOTE_PROJECTION_PAGE_DEPTH",
            () => section.ToMarkdown(new OneNoteMarkdownOptions {
                IncludeConflictPages = true,
                MaxPageRelationshipDepth = 1
            }));

        var parent = new OneNoteOutline();
        parent.Children.Add(new OneNoteParagraph());
        AssertProjectionError(
            "ONENOTE_PROJECTION_CONTENT_DEPTH",
            () => new OneNoteSection {
                Pages = { new OneNotePage { Outlines = { parent } } }
            }.ToMarkdown(new OneNoteMarkdownOptions { MaxContentDepth = 1 }));
    }

    [Fact]
    public void ProjectionRejectsSharedModelInstancesBeforeRepeatedExpansion() {
        var sharedContent = new OneNoteParagraph();
        var contentPage = new OneNotePage { Title = "Content" };
        contentPage.DirectContent.Add(sharedContent);
        contentPage.DirectContent.Add(sharedContent);
        var contentSection = new OneNoteSection { Pages = { contentPage } };
        AssertProjectionError("ONENOTE_PROJECTION_SHARED_CONTENT", () => contentSection.ToMarkdown());

        var sharedPage = new OneNotePage { Title = "Page" };
        var pageSection = new OneNoteSection { Pages = { sharedPage, sharedPage } };
        AssertProjectionError("ONENOTE_PROJECTION_SHARED_PAGE", () => pageSection.ToMarkdown());

        var sharedGroup = new OneNoteSectionGroup { Name = "Group" };
        var notebook = new OneNoteNotebook { SectionGroups = { sharedGroup, sharedGroup } };
        AssertProjectionError("ONENOTE_PROJECTION_SHARED_GROUP", () => notebook.ToMarkdown());
    }

    [Fact]
    public void ProjectionDepthOptionsEnforceTheHardTraversalCeiling() {
        int invalid = OneNoteWriterOptions.MaximumTraversalDepth + 1;
        var section = new OneNoteSection();

        Assert.Throws<ArgumentOutOfRangeException>(() => section.ToMarkdown(new OneNoteMarkdownOptions { MaxSectionGroupDepth = invalid }));
        Assert.Throws<ArgumentOutOfRangeException>(() => section.ToMarkdown(new OneNoteMarkdownOptions { MaxPageRelationshipDepth = invalid }));
        Assert.Throws<ArgumentOutOfRangeException>(() => section.ToMarkdown(new OneNoteMarkdownOptions { MaxContentDepth = invalid }));
    }

    [Fact]
    public void MarkdownProjectionKeepsLiteralLineStartsInsideTheirSourceBlocks() {
        var section = new OneNoteSection { Name = "Projection" };
        var page = new OneNotePage { Title = "Title\n# literal heading" };
        page.DirectContent.Add(Paragraph("Text\n> literal quote\n- literal item\n1. literal ordered item"));
        section.Pages.Add(page);

        string markdown = section.ToMarkdown();
        string html = section.ToHtmlDocument(htmlOptions: new HtmlOptions { AssetMode = AssetMode.Offline });

        Assert.Contains("## Title<br># literal heading", markdown);
        Assert.Contains("Text<br>&gt; literal quote<br>- literal item<br>1. literal ordered item", markdown);
        Assert.DoesNotContain("\n# literal heading", markdown);
        Assert.DoesNotContain("<blockquote", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<li", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownAndHtmlProjectionEscapeSourceHtmlAsLiteralText() {
        var section = new OneNoteSection { Name = "Projection" };
        var page = new OneNotePage { Title = "<script>alert('title')</script>" };
        page.DirectContent.Add(Paragraph("<img src=x onerror=alert('body')>"));
        section.Pages.Add(page);

        string markdown = section.ToMarkdown();
        string html = section.ToHtmlDocument(htmlOptions: new HtmlOptions { AssetMode = AssetMode.Offline });

        Assert.Contains("&lt;script&gt;", markdown, StringComparison.Ordinal);
        Assert.Contains("&lt;img", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<img src=x", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&lt;script&gt;", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&lt;img", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownDestinationsEncodeControlsAndHtmlSignificantCharacters() {
        const string unsafeDestination = "https://example.invalid/path\n<script>alert(1)</script>";
        var section = new OneNoteSection { Name = "Projection" };
        var page = new OneNotePage { Title = "Unsafe links" };
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = "link", Hyperlink = unsafeDestination });
        page.DirectContent.Add(paragraph);
        page.DirectContent.Add(new OneNoteImage {
            AltText = "image",
            Hyperlink = unsafeDestination,
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1 })
        });
        section.Pages.Add(page);
        var projectionOptions = new OneNoteMarkdownOptions { AssetUriResolver = _ => unsafeDestination };

        string markdown = section.ToMarkdown(projectionOptions);
        string html = section.ToHtmlDocument(
            projectionOptions,
            new HtmlOptions { AssetMode = AssetMode.Offline });

        Assert.Contains("https://example.invalid/path%0A%3Cscript%3Ealert%281%29%3C/script%3E", markdown);
        Assert.DoesNotContain("\n<script", markdown, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("%0A%3Cscript%3E", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MathProjectionUsesAFenceThatCannotBeClosedBySourceContent() {
        var section = new OneNoteSection { Name = "Projection" };
        var page = new OneNotePage { Title = "Unsafe math" };
        page.DirectContent.Add(new OneNoteMath {
            Latex = "```\n~~~\n<script>alert('math')</script>"
        });
        section.Pages.Add(page);

        string markdown = section.ToMarkdown();
        string html = section.ToHtmlDocument(htmlOptions: new HtmlOptions { AssetMode = AssetMode.Offline });

        Assert.Contains("````math", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain("<script", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("&lt;script&gt;", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownTableCellEscapesPipeExactlyOnce() {
        var section = new OneNoteSection { Name = "Projection" };
        var page = new OneNotePage { Title = "Table" };
        var table = new OneNoteTable();
        var row = new OneNoteTableRow();
        row.Cells.Add(Cell("Left | Right"));
        table.Rows.Add(row);
        page.DirectContent.Add(table);
        section.Pages.Add(page);

        string markdown = section.ToMarkdown();

        Assert.Contains(@"Left \| Right", markdown, StringComparison.Ordinal);
        Assert.DoesNotContain(@"\\|", markdown, StringComparison.Ordinal);
    }

    private static OneNoteSection SectionWithPage(string sectionName, string pageTitle) {
        var section = new OneNoteSection { Name = sectionName };
        section.Pages.Add(new OneNotePage { Title = pageTitle });
        return section;
    }

    private static void AssertProjectionError(string code, Action action) {
        OneNoteFormatException exception = Assert.Throws<OneNoteFormatException>(action);
        Assert.Equal(code, exception.Code);
    }

    private static OneNoteNotebook CreateNotebook() {
        var notebook = new OneNoteNotebook { Name = "Offline notebook" };
        var group = new OneNoteSectionGroup { Name = "Group A" };
        var section = new OneNoteSection { Name = "Section A" };
        var page = new OneNotePage { Title = "Current page" };

        var paragraph = new OneNoteParagraph();
        var bold = new OneNoteTextRun { Text = "Bold" };
        bold.Style.Bold = true;
        paragraph.Runs.Add(bold);
        paragraph.Runs.Add(new OneNoteTextRun { Text = " " });
        paragraph.Runs.Add(new OneNoteTextRun { Text = "link", Hyperlink = "https://example.com/a b" });
        page.DirectContent.Add(paragraph);

        var table = new OneNoteTable();
        var row = new OneNoteTableRow();
        row.Cells.Add(Cell("Left"));
        row.Cells.Add(Cell("Right"));
        table.Rows.Add(row);
        page.DirectContent.Add(table);

        var listItem = new OneNoteParagraph { List = new OneNoteListInfo { Level = 0 } };
        listItem.Runs.Add(new OneNoteTextRun { Text = "Item" });
        page.DirectContent.Add(listItem);
        page.DirectContent.Add(new OneNoteImage {
            FileName = "diagram.png",
            AltText = "Diagram",
            MediaType = "image/png",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 1, 2, 3 })
        });
        page.DirectContent.Add(new OneNoteEmbeddedFile {
            FileName = "sample.zip",
            Payload = OneNoteBinaryPayload.FromBytes(new byte[] { 4, 5, 6 })
        });
        page.DirectContent.Add(new OneNoteMath { Text = "x^2", Latex = "x^2" });

        var conflict = new OneNotePage { Title = "Conflict copy", IsConflictPage = true };
        conflict.DirectContent.Add(Paragraph("Conflict body"));
        conflict.VersionHistory.Add(new OneNotePage { Title = "Conflict history", IsVersionHistoryPage = true });
        page.ConflictPages.Add(conflict);
        var version = new OneNotePage { Title = "Historical copy", IsVersionHistoryPage = true };
        version.DirectContent.Add(Paragraph("Historical body"));
        version.ConflictPages.Add(new OneNotePage { Title = "Historical conflict", IsConflictPage = true });
        page.VersionHistory.Add(version);

        section.Pages.Add(page);
        group.Sections.Add(section);
        notebook.SectionGroups.Add(group);
        return notebook;
    }

    private static OneNoteTableCell Cell(string text) {
        var cell = new OneNoteTableCell();
        cell.Content.Add(Paragraph(text));
        return cell;
    }

    private static OneNoteParagraph Paragraph(string text) {
        var paragraph = new OneNoteParagraph();
        paragraph.Runs.Add(new OneNoteTextRun { Text = text });
        return paragraph;
    }
}
