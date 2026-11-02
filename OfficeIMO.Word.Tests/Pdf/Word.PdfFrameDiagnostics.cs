using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("body", false, false)]
    [InlineData("body", true, false)]
    [InlineData("table", false, false)]
    [InlineData("footnote", false, false)]
    [InlineData("endnote", false, false)]
    [InlineData("header", false, false)]
    [InlineData("footer", false, false)]
    [InlineData("table", false, true)]
    [InlineData("footnote", false, true)]
    [InlineData("endnote", false, true)]
    [InlineData("header", false, true)]
    [InlineData("footer", false, true)]
    public void SaveAsPdf_ParagraphFramesReportApproximationAndPreserveSource(string story, bool inherited, bool breakOnly) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("BODY");
        WordParagraph paragraph = AddFrameDiagnosticParagraph(document, story);
        paragraph.Text = breakOnly ? string.Empty : "FRAMEDMARKER";
        if (breakOnly) paragraph.AddBreak(WordBreakType.Page);
        SetFrameDiagnosticProperties(document, paragraph, inherited);
        Assert.Empty(document.ValidateDocument());
        string paragraphXml = paragraph._paragraph.OuterXml;
        string stylesXml = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml;
        var options = new WordPdfExportOptions();

        for (int index = 0; index < 2; index++) {
            var result = document.ConvertToPdfWithEvidence(options);
            PdfConversionWarning warning = Assert.Single(result.Warnings,
                item => item.Code == "native-paragraph-frame-flattened");
            Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, warning.LossKind);
            Assert.True(result.HasLoss);
            Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
            using var output = new MemoryStream();
            var saved = result.Save(output);
            Assert.True(saved.Succeeded);
            Assert.True(saved.HasLoss);
            Assert.Single(saved.Warnings, item => item.Code == "native-paragraph-frame-flattened");
            byte[] savedPdf = output.ToArray();
            string allText = PdfReadDocument.Open(savedPdf,
                new PdfLoadOptions { IncludeArtifactText = true }).ExtractText();
            Assert.Contains("BODY", allText, StringComparison.Ordinal);
            if (!breakOnly) Assert.Contains("FRAMEDMARKER", allText, StringComparison.Ordinal);
            if (!breakOnly && (story is "header" or "footer"))
                Assert.DoesNotContain("FRAMEDMARKER", PdfReadDocument.Open(savedPdf).ExtractText(), StringComparison.Ordinal);
            Assert.Equal(paragraphXml, paragraph._paragraph.OuterXml);
            Assert.Equal(stylesXml, document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml);
        }
    }

    [Fact]
    public void SaveAsPdf_OrdinaryParagraphsDoNotReportFrameLoss() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("PLAINMARKER");
        var result = document.ConvertToPdfWithEvidence();
        Assert.DoesNotContain(result.Warnings, item => item.Code == "native-paragraph-frame-flattened");
        Assert.Contains("PLAINMARKER", PdfReadDocument.Open(result.ToBytes()).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, true)]
    public void SaveAsPdf_BreakOnlyParagraphFramesReportApproximation(bool columnBreak, bool inherited) {
        using WordDocument document = CreateFrameFlowBreakDocument(columnBreak, inherited, out WordParagraph paragraph);
        string sourceXml = paragraph._paragraph.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(CreateFrameFlowBreakPdfOptions());

        PdfConversionWarning warning = Assert.Single(result.Warnings, item => item.Code == "native-paragraph-frame-flattened");
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Approximation, warning.LossKind);
        Assert.True(result.HasLoss);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
        string text = PdfReadDocument.Open(result.ToBytes()).ExtractText();
        Assert.Contains("BEFORE", text, StringComparison.Ordinal);
        Assert.Contains("AFTER", text, StringComparison.Ordinal);
        Assert.Equal(sourceXml, paragraph._paragraph.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_BreakOnlyFrameRequireNoLossRejectsApproximation() {
        using WordDocument document = CreateFrameFlowBreakDocument(false, false, out _);

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(CreateFrameFlowBreakPdfOptions());

        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
        Assert.Throws<InvalidOperationException>(() => result.Report.RequireNoLoss());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    public void SaveAsPdf_SplitParagraphFrameReportsOnceAndPreservesContent(bool columnBreak, bool trailingLegacyMark) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = trailingLegacyMark
            ? WordCompatibilityMode.Word2003 : WordCompatibilityMode.Word2013;
        document.CompatibilitySettings.SplitPageBreakAndParagraphMark = trailingLegacyMark;
        if (columnBreak) document.Sections[0].ColumnCount = 2;
        WordParagraph paragraph = document.AddParagraph("FRAMEBEFORE");
        paragraph.AddBreak(columnBreak ? WordBreakType.Column : WordBreakType.Page);
        if (!trailingLegacyMark) paragraph.AddText("FRAMEAFTER");
        document.AddParagraph("TAIL");
        SetFrameDiagnosticProperties(document, paragraph, inherited: false);
        string sourceXml = paragraph._paragraph.OuterXml;

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions { IncludePageNumbers = false });

        Assert.Single(result.Warnings, item => item.Code == "native-paragraph-frame-flattened");
        string text = PdfReadDocument.Open(result.ToBytes()).ExtractText();
        Assert.Equal(1, text.Split(new[] { "FRAMEBEFORE" }, StringSplitOptions.None).Length - 1);
        if (!trailingLegacyMark) Assert.Equal(1, text.Split(new[] { "FRAMEAFTER" }, StringSplitOptions.None).Length - 1);
        Assert.Contains("TAIL", text, StringComparison.Ordinal);
        Assert.Equal(sourceXml, paragraph._paragraph.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_BreakOnlyParagraphWithoutFrameRemainsLossless(bool columnBreak) {
        using WordDocument document = CreateFrameFlowBreakDocument(columnBreak, false, out WordParagraph paragraph);
        paragraph._paragraph.ParagraphProperties!.RemoveAllChildren<W.FrameProperties>();

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(CreateFrameFlowBreakPdfOptions());

        Assert.DoesNotContain(result.Warnings, item => item.Code == "native-paragraph-frame-flattened");
        result.RequireNoLoss();
        string text = PdfReadDocument.Open(result.ToBytes()).ExtractText();
        Assert.Contains("BEFORE", text, StringComparison.Ordinal);
        Assert.Contains("AFTER", text, StringComparison.Ordinal);
    }

    private static WordDocument CreateFrameFlowBreakDocument(bool columnBreak, bool inherited, out WordParagraph paragraph) {
        WordDocument document = WordDocument.Create();
        document.Settings.FontFamily = ManagedTextShapingTestAssets.FamilyName;
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
        if (columnBreak) document.Sections[0].ColumnCount = 2;
        document.AddParagraph("BEFORE");
        paragraph = document.AddParagraph();
        paragraph.AddBreak(columnBreak ? WordBreakType.Column : WordBreakType.Page);
        document.AddParagraph("AFTER");
        SetFrameDiagnosticProperties(document, paragraph, inherited);
        return document;
    }

    private static WordPdfExportOptions CreateFrameFlowBreakPdfOptions() {
        var pdfOptions = new PdfCreateOptions();
        pdfOptions.RegisterNamedFontFamily(new PdfEmbeddedFontFamily(
            ManagedTextShapingTestAssets.FamilyName,
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B', 'E', 'F', 'O', 'R', 'T')));
        return new WordPdfExportOptions {
            IncludePageNumbers = false,
            ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PdfOptions = pdfOptions
        };
    }

    private static void SetFrameDiagnosticProperties(WordDocument document, WordParagraph paragraph, bool inherited) {
        var frame = new W.FrameProperties();
        frame.SetAttribute(new OpenXmlAttribute("w", "w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "3600"));
        frame.SetAttribute(new OpenXmlAttribute("w", "x", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "1440"));
        frame.SetAttribute(new OpenXmlAttribute("w", "y", "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "720"));
        if (inherited) {
            document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.StyleParagraphProperties(frame)) {
                    Type = W.StyleValues.Paragraph, StyleId = "PositionedFrame"
                });
            paragraph.SetStyleId("PositionedFrame");
        } else {
            paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.AddChild(frame, true);
        }
    }

    private static WordParagraph AddFrameDiagnosticParagraph(WordDocument document, string story) {
        switch (story) {
            case "body": return document.AddParagraph();
            case "table": return document.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0];
            case "header": return document.HeaderDefaultOrCreate.AddParagraph();
            case "footer": return document.FooterDefaultOrCreate.AddParagraph();
            default:
                WordParagraph body = document.AddParagraph("REFERENCE");
                bool endnote = story == "endnote";
                if (endnote) body.AddEndNote("NOTE");
                else body.AddFootNote("NOTE");
                var main = document.OpenXmlDocument.MainDocumentPart!;
                W.FootnoteEndnoteType note = endnote
                    ? Assert.Single(main.EndnotesPart!.Endnotes!.Elements<W.Endnote>(), item => item.Id?.Value ==
                        Assert.Single(body._paragraph.Descendants<W.EndnoteReference>()).Id!.Value)
                    : Assert.Single(main.FootnotesPart!.Footnotes!.Elements<W.Footnote>(), item => item.Id?.Value ==
                        Assert.Single(body._paragraph.Descendants<W.FootnoteReference>()).Id!.Value);
                return new WordParagraph(document, Assert.Single(note.Elements<W.Paragraph>()));
        }
    }
}
