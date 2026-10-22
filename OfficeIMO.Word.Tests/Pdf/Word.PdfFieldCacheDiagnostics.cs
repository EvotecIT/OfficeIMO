using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("complex-body", "body", true, " LISTNUM CustomHeadings \\l 5 ")]
    [InlineData("simple-body", "body", false, " LISTNUM CustomHeadings \\l 5 ")]
    [InlineData("simple-table", "table", false, " REF MissingBookmark ")]
    [InlineData("complex-header", "header", true, " DATE ")]
    [InlineData("simple-footer", "footer", false, " NUMWORDS ")]
    [InlineData("body-page-field", "body", true, " PAGE ")]
    public void SaveAsPdf_FieldCache_MissingDisplayResultReportsOmissionWithoutChangingSource(
        string caseId, string story, bool complex, string instruction) {
        _ = caseId;
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        document.AddParagraph("BODY");
        WordParagraph paragraph = AddFieldCacheDiagnosticParagraph(document, story);
        paragraph._paragraph.Append(new W.Run(new W.Text("BEFORE")));
        AppendDiagnosticField(paragraph, complex, instruction, cache: "missing");
        paragraph._paragraph.Append(new W.Run(new W.Text("AFTER")));
        Assert.Empty(document.ValidateDocument());
        string sourceXml = paragraph._paragraph.OuterXml;

        var result = document.ToPdfDocumentResult(FieldCacheDiagnosticOptions());
        var warning = Assert.Single(result.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.Equal(PdfConversionWarningSeverity.Warning, warning.Severity);
        Assert.True(result.HasLoss);
        Assert.False(result.Report.HasErrors);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("BEFOREAFTER", visible, StringComparison.Ordinal);
        Assert.DoesNotContain(instruction.Trim(), visible, StringComparison.Ordinal);
        Assert.Equal(sourceXml, paragraph._paragraph.OuterXml);
        Assert.Empty(Assert.Single(document.InspectFields()).ResultText);
    }

    [Theory]
    [InlineData("complex-empty", true, "empty")]
    [InlineData("simple-empty", false, "empty")]
    [InlineData("complex-stored", true, "stored")]
    [InlineData("simple-stored", false, "stored")]
    public void SaveAsPdf_FieldCache_ExplicitEmptyAndStoredResultsDoNotReportMissingCache(
        string caseId, bool complex, string cache) {
        _ = caseId;
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        WordParagraph paragraph = document.AddParagraph("BEFORE");
        AppendDiagnosticField(paragraph, complex, " LISTNUM CustomHeadings \\l 5 ", cache);
        paragraph._paragraph.Append(new W.Run(new W.Text("AFTER")));
        Assert.Empty(document.ValidateDocument());
        var result = document.ToPdfDocumentResult(FieldCacheDiagnosticOptions());
        Assert.DoesNotContain(result.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        result.RequireNoLoss();
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains(cache == "stored" ? "BEFORESTOREDRESULTAFTER" : "BEFOREAFTER", visible, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("direct-hidden", "direct", false)]
    [InlineData("character-hidden", "character", false)]
    [InlineData("paragraph-hidden", "paragraph", false)]
    [InlineData("document-hidden", "document", false)]
    [InlineData("visible-override", "character", true)]
    public void SaveAsPdf_FieldCache_MissingResultDiagnosticUsesEffectiveVisibility(
        string caseId, string hiddenSource, bool visibleOverride) {
        _ = caseId;
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        document.AddParagraph("BODY");
        W.Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        styles.Append(new W.Style(new W.StyleRunProperties(new W.Vanish())) {
            Type = hiddenSource == "paragraph" ? W.StyleValues.Paragraph : W.StyleValues.Character,
            StyleId = "HiddenMissingField"
        });
        if (hiddenSource == "document") styles.DocDefaults!.RunPropertiesDefault!.RunPropertiesBaseStyle!
            .AddChild(new W.Vanish(), true);
        WordParagraph paragraph = document.AddParagraph();
        if (hiddenSource == "paragraph") paragraph.SetStyleId("HiddenMissingField");
        var properties = new W.RunProperties();
        if (hiddenSource == "direct") properties.AddChild(new W.Vanish(), true);
        if (hiddenSource == "character") properties.AddChild(new W.RunStyle { Val = "HiddenMissingField" }, true);
        if (visibleOverride) properties.AddChild(new W.Vanish { Val = false }, true);
        AppendDiagnosticField(paragraph, complex: true, " LISTNUM CustomHeadings \\l 5 ", "missing", properties);
        Assert.Empty(document.ValidateDocument());

        var result = document.ToPdfDocumentResult(FieldCacheDiagnosticOptions());
        Assert.Equal(visibleOverride ? 1 : 0, result.Warnings.Count(item => item.Code == "NativeFieldResultUnavailable"));
        if (visibleOverride) Assert.True(result.HasLoss);
        else result.RequireNoLoss();
    }

    [Theory]
    [InlineData("set-variable", " SET Bookmark 1 ")]
    [InlineData("ask-input", " ASK Bookmark \"Question\" ")]
    [InlineData("toc-entry", " TC \"Entry\" ")]
    [InlineData("index-entry", " XE \"Entry\" ")]
    public void SaveAsPdf_FieldCache_NonDisplayFieldsDoNotRequireDisplayResult(string caseId, string instruction) {
        _ = caseId;
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        WordParagraph paragraph = document.AddParagraph("BEFORE");
        AppendDiagnosticField(paragraph, complex: true, instruction, "missing");
        paragraph._paragraph.Append(new W.Run(new W.Text("AFTER")));
        Assert.Empty(document.ValidateDocument());
        var result = document.ToPdfDocumentResult(FieldCacheDiagnosticOptions());
        Assert.DoesNotContain(result.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        result.RequireNoLoss();
    }

    [Theory]
    [InlineData("simple-header", false, false)]
    [InlineData("complex-footer", true, true)]
    public void SaveAsPdf_FieldCache_CalculatedRunningCountersDoNotRequireStoredResults(string caseId, bool complex, bool footer) {
        _ = caseId;
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        document.AddParagraph("BODY");
        WordParagraph paragraph = AddFieldCacheDiagnosticParagraph(document, footer ? "footer" : "header");
        paragraph._paragraph.Append(new W.Run(new W.Text("RUNNING")));
        foreach (string instruction in new[] { " PAGE ", " NUMPAGES ", " SECTIONPAGES " }) {
            AppendDiagnosticField(paragraph, complex, instruction, "missing");
            paragraph._paragraph.Append(new W.Run(new W.Text("END")));
        }
        Assert.Empty(document.ValidateDocument());
        var result = document.ToPdfDocumentResult(FieldCacheDiagnosticOptions());
        Assert.DoesNotContain(result.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        result.RequireNoLoss();
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("RUNNING1END1END1END", visible, StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_FieldCache_InactiveStoredRunningFieldDoesNotReportMissingCache() {
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        document.DifferentFirstPage = true;
        document.AddHeadersAndFooters();
        WordHeader target = RequireSectionHeader(document, 0, W.HeaderFooterValues.First);
        WordParagraph paragraph = target.AddParagraph();
        AppendDiagnosticField(paragraph, complex: true, " LISTNUM CustomHeadings \\l 5 ", "missing");
        document.AddParagraph("BODY");
        document.Sections[0]._sectionProperties.RemoveAllChildren<W.TitlePage>();
        Assert.Single(document.InspectFields());
        var result = document.ToPdfDocumentResult(FieldCacheDiagnosticOptions());
        Assert.DoesNotContain(result.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        result.RequireNoLoss();
    }

    [Fact]
    public void SaveAsPdf_FieldCache_InstructionNestedFieldDoesNotReportVisibleResultLoss() {
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        WordParagraph paragraph = document.AddParagraph("BEFORE");
        paragraph._paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" QUOTE ")));
        AppendDiagnosticField(paragraph, complex: true, " LISTNUM CustomHeadings \\l 5 ", "missing");
        paragraph._paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            new W.Run(new W.Text("STOREDRESULT")),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }),
            new W.Run(new W.Text("AFTER")));
        Assert.Empty(document.ValidateDocument());
        var result = document.ToPdfDocumentResult(FieldCacheDiagnosticOptions());
        Assert.DoesNotContain(result.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        result.RequireNoLoss();
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        Assert.Contains("BEFORESTOREDRESULTAFTER", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)), StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_FieldCache_ReusedOptionsRebuildsInventoryAndKeepsEarlierReports() {
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        WordParagraph paragraph = document.AddParagraph("BEFORE");
        AppendDiagnosticField(paragraph, complex: true, " LISTNUM CustomHeadings \\l 5 ", "missing");
        paragraph._paragraph.Append(new W.Run(new W.Text("AFTER")));
        WordToPdfOptions options = FieldCacheDiagnosticOptions();
        var first = document.ToPdfDocumentResult(options);
        var repeated = document.ToPdfDocumentResult(options);
        Assert.Single(first.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        Assert.Single(repeated.Warnings, item => item.Code == "NativeFieldResultUnavailable");

        W.Run end = Assert.Single(paragraph._paragraph.Elements<W.Run>(), run =>
            run.Elements<W.FieldChar>().Any(marker => marker.FieldCharType?.Value == W.FieldCharValues.End));
        end.InsertBeforeSelf(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }));
        end.InsertBeforeSelf(new W.Run(new W.Text("STOREDRESULT")));
        Assert.Empty(document.ValidateDocument());
        var afterEdit = document.ToPdfDocumentResult(options);
        Assert.DoesNotContain(afterEdit.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        afterEdit.RequireNoLoss();
        Assert.Single(first.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        Assert.Single(repeated.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        using var pdf = PdfPigDocument.Open(afterEdit.ToBytes());
        Assert.Contains("BEFORESTOREDRESULTAFTER", string.Concat(pdf.GetPage(1).Letters.Select(letter => letter.Value)), StringComparison.Ordinal);
    }

    [Fact]
    public void SaveAsPdf_FieldCache_ComplexFieldAcrossParagraphsReportsOneOmission() {
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        WordParagraph start = document.AddParagraph("BEFORE");
        start._paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode(" LISTNUM CustomHeadings \\l 5 ")));
        WordParagraph end = document.AddParagraph();
        end._paragraph.Append(new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }),
            new W.Run(new W.Text("AFTER")));
        Assert.Empty(document.ValidateDocument());
        string[] xml = { start._paragraph.OuterXml, end._paragraph.OuterXml };
        var result = document.ToPdfDocumentResult(FieldCacheDiagnosticOptions());
        Assert.Single(result.Warnings, item => item.Code == "NativeFieldResultUnavailable");
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("BEFORE", visible, StringComparison.Ordinal);
        Assert.Contains("AFTER", visible, StringComparison.Ordinal);
        Assert.Equal(xml, new[] { start._paragraph.OuterXml, end._paragraph.OuterXml });
    }

    [Theory]
    [InlineData("footnote-missing", false, "missing", false)]
    [InlineData("endnote-missing", true, "missing", false)]
    [InlineData("footnote-hidden", false, "missing", true)]
    [InlineData("endnote-hidden", true, "missing", true)]
    [InlineData("footnote-empty", false, "empty", false)]
    [InlineData("endnote-empty", true, "empty", false)]
    public void SaveAsPdf_FieldCache_NoteFallbackReportsVisibleMissingResultsOnce(
        string caseId, bool endnote, string cache, bool hiddenField) {
        _ = caseId;
        using WordDocument document = CreateFieldCacheDiagnosticDocument();
        WordParagraph body = document.AddParagraph("BODY");
        if (endnote) body.AddEndNote("NOTEBEFORE");
        else body.AddFootNote("NOTEBEFORE");
        var main = document._wordprocessingDocument.MainDocumentPart!;
        DocumentFormat.OpenXml.OpenXmlElement reference = endnote
            ? Assert.Single(body._paragraph.Descendants<W.EndnoteReference>())
            : Assert.Single(body._paragraph.Descendants<W.FootnoteReference>());
        W.FootnoteEndnoteType note = endnote
            ? Assert.Single(main.EndnotesPart!.Endnotes!.Elements<W.Endnote>(), item =>
                item.Id?.Value == ((W.EndnoteReference)reference).Id!.Value)
            : Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>(), item =>
                item.Id?.Value == ((W.FootnoteReference)reference).Id!.Value);
        var paragraph = new WordParagraph(document, Assert.Single(note.Elements<W.Paragraph>()));
        if (hiddenField) {
            main.StyleDefinitionsPart!.Styles!.Append(new W.Style(new W.StyleRunProperties(new W.Vanish())) {
                Type = W.StyleValues.Paragraph, StyleId = "HiddenMissingNoteField"
            });
            paragraph.SetStyleId("HiddenMissingNoteField");
            foreach (W.Run run in paragraph._paragraph.Elements<W.Run>()) {
                run.RunProperties ??= new W.RunProperties();
                run.RunProperties.AddChild(new W.Vanish { Val = false }, true);
            }
        }
        AppendDiagnosticField(paragraph, complex: !endnote, " LISTNUM CustomHeadings \\l 5 ", cache);
        paragraph._paragraph.Append(new W.Run(new W.RunProperties(new W.Vanish { Val = false }), new W.Text("NOTEAFTER")));
        // Repeated references retain one logical note and one field diagnostic.
        document.AddParagraph("REPEATEDBODY")._paragraph.Append(new W.Run(reference.CloneNode(true)));
        Assert.Empty(document.ValidateDocument());
        string sourceXml = note.OuterXml;

        var result = document.ToPdfDocumentResult(FieldCacheDiagnosticOptions());
        var warnings = result.Warnings.Where(item => item.Code == "NativeFieldResultUnavailable").ToArray();
        if (cache == "missing" && !hiddenField) {
            var warning = Assert.Single(warnings);
            Assert.Equal(OfficeIMO.OfficeConversionLossKind.Omission, warning.LossKind);
            Assert.StartsWith(endnote ? "endnote " : "footnote ", warning.Source, StringComparison.Ordinal);
        } else Assert.Empty(warnings);
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        string visible = string.Concat(pdf.GetPages().SelectMany(page => page.Letters).Select(letter => letter.Value));
        Assert.Contains("NOTEBEFORE", visible, StringComparison.Ordinal);
        Assert.Contains("NOTEAFTER", visible, StringComparison.Ordinal);
        Assert.Contains("REPEATEDBODY", visible, StringComparison.Ordinal);
        Assert.Equal(sourceXml, note.OuterXml);
    }

    private static WordDocument CreateFieldCacheDiagnosticDocument() {
        WordDocument document = CreateJoinedParagraphDocument();
        // Use native PDF fonts so the no-loss gate isolates field results rather
        // than the deliberately disabled source-font embedding policy.
        foreach (W.RunFonts fonts in document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Descendants<W.RunFonts>()) {
            fonts.Ascii = "Helvetica";
            fonts.HighAnsi = "Helvetica";
            fonts.EastAsia = "Helvetica";
            fonts.ComplexScript = "Helvetica";
            fonts.AsciiTheme = null;
            fonts.HighAnsiTheme = null;
            fonts.EastAsiaTheme = null;
            fonts.ComplexScriptTheme = null;
        }
        return document;
    }

    private static WordParagraph AddFieldCacheDiagnosticParagraph(WordDocument document, string story) => story switch {
        "body" => document.AddParagraph(),
        "table" => document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0],
        "header" => document.HeaderDefaultOrCreate.AddParagraph(),
        "footer" => document.FooterDefaultOrCreate.AddParagraph(),
        _ => throw new ArgumentOutOfRangeException(nameof(story))
    };

    private static void AppendDiagnosticField(WordParagraph paragraph, bool complex, string instruction, string cache,
        W.RunProperties? properties = null) {
        W.Run Run(params DocumentFormat.OpenXml.OpenXmlElement[] content) {
            var run = new W.Run();
            if (properties != null) run.Append(properties.CloneNode(true));
            run.Append(content);
            return run;
        }
        if (complex) {
            paragraph._paragraph.Append(Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
                Run(new W.FieldCode(instruction)));
            if (cache != "missing") paragraph._paragraph.Append(Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }));
            if (cache == "stored") paragraph._paragraph.Append(Run(new W.Text("STOREDRESULT")));
            paragraph._paragraph.Append(Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        } else {
            var field = new W.SimpleField { Instruction = instruction };
            if (cache != "missing") field.Append(Run(new W.Text(cache == "stored" ? "STOREDRESULT" : string.Empty)));
            paragraph._paragraph.Append(field);
        }
    }

    private static WordToPdfOptions FieldCacheDiagnosticOptions() => new() {
        IncludePageNumbers = false,
        ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
    };
}
