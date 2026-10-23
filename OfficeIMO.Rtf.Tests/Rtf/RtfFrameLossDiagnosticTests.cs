using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.Rtf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfFrameLossDiagnosticTests {
    [Theory]
    [InlineData("body")]
    [InlineData("table")]
    [InlineData("footnote")]
    [InlineData("endnote")]
    [InlineData("header")]
    [InlineData("footer")]
    public void EffectiveFrameLossIsReportedByDirectAndBridgeRoutesWithoutChangingSource(string story) {
        RtfDocument source = RtfDocument.Create();
        source.AddStyle(7, "Positioned").Frame.SetSize(3600).SetPosition(horizontalTwips: 1440, verticalTwips: 720);
        source.AddParagraph("BODY");
        RtfParagraph framed = AddParagraph(source, story);
        framed.SetStyle(7).AddText("FRAMEDMARKER");
        string original = source.ToRtf();

        PdfCore.PdfDocumentConversionResult direct = source.ToPdfDocumentResult();
        PdfCore.PdfConversionWarning directWarning = Assert.Single(direct.Warnings,
            warning => warning.Code == "ParagraphFrameFlattened");
        Assert.Equal(OfficeConversionLossKind.Approximation, directWarning.LossKind);
        Assert.True(direct.HasLoss);
        Assert.Throws<InvalidOperationException>(() => direct.RequireNoLoss());
        Assert.Contains("FRAMEDMARKER", PdfCore.PdfReadDocument.Open(direct.ToBytes()).ExtractText(), StringComparison.Ordinal);

        RtfConversionResult<WordDocument> bridge = source.ToWordDocumentResult();
        using WordDocument word = bridge.Value;
        if (story is "header" or "footer") {
            WordHeaderFooter runningStory = story == "header"
                ? Assert.IsType<WordHeader>(word.Header?.Default)
                : Assert.IsType<WordFooter>(word.Footer?.Default);
            Assert.Contains("FRAMEDMARKER", string.Concat(runningStory.Paragraphs.Select(item => item.Text)), StringComparison.Ordinal);
            var section = Assert.Single(word.Sections);
            Assert.False(section.DifferentFirstPage);
            string relationshipId = story == "header"
                ? Assert.Single(section._sectionProperties.Elements<W.HeaderReference>(),
                    reference => reference.Type?.Value == W.HeaderFooterValues.Default).Id!.Value!
                : Assert.Single(section._sectionProperties.Elements<W.FooterReference>(),
                    reference => reference.Type?.Value == W.HeaderFooterValues.Default).Id!.Value!;
            var part = word.OpenXmlDocument.MainDocumentPart!.GetPartById(relationshipId);
            string nativeText = part switch {
                DocumentFormat.OpenXml.Packaging.HeaderPart header => header.Header!.InnerText,
                DocumentFormat.OpenXml.Packaging.FooterPart footer => footer.Footer!.InnerText,
                _ => throw new InvalidOperationException("The running-story reference must resolve to its native part.")
            };
            Assert.Contains("FRAMEDMARKER", nativeText, StringComparison.Ordinal);
        }
        RtfConversionDiagnostic bridgeWarning = Assert.Single(bridge.Report.Diagnostics,
            diagnostic => diagnostic.Code == "RtfWordParagraphFramesFlattened");
        Assert.Equal(RtfConversionAction.Flattened, bridgeWarning.Action);
        Assert.Equal(1, bridgeWarning.Count);
        Assert.True(bridge.HasLoss);
        Assert.Throws<RtfConversionLossException>(() => bridge.RequireNoLoss());
        using var output = new MemoryStream();
        PdfCore.PdfSaveResult saved = word.ToPdfDocumentResult().WithSourceConversionReport(bridge.Report).Save(output);
        Assert.True(saved.Succeeded);
        Assert.True(saved.HasLoss);
        Assert.Contains(saved.FidelityDiagnostics, diagnostic => diagnostic.Code == "RtfWordParagraphFramesFlattened" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.Throws<RtfConversionLossException>(() => saved.RequireNoLoss());
        byte[] savedPdf = output.ToArray();
        Assert.Contains("FRAMEDMARKER", PdfCore.PdfReadDocument.Open(savedPdf,
            new PdfCore.PdfLoadOptions { IncludeArtifactText = true }).ExtractText(), StringComparison.Ordinal);
        if (story is "header" or "footer") {
            // Running stories are tagged pagination artifacts. Ordinary text
            // extraction excludes them; that is not evidence of missing paint.
            Assert.DoesNotContain("FRAMEDMARKER", PdfCore.PdfReadDocument.Open(savedPdf).ExtractText(), StringComparison.Ordinal);
        }
        Assert.Equal(original, source.ToRtf());
        Assert.False(framed.Frame.HasAnyValue);
    }

    [Fact]
    public void FrameWarningsAreScopedToEachConversionAndDirectSaveReport() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("FRAMEDMARKER").Frame.WidthTwips = 3600;
        var options = new RtfToPdfOptions();
        for (int index = 0; index < 2; index++) {
            using var output = new MemoryStream();
            PdfCore.PdfSaveResult saved = source.SaveAsPdf(output, options);
            Assert.Single(saved.Warnings, warning => warning.Code == "ParagraphFrameFlattened");
            Assert.True(saved.HasLoss);
            Assert.Contains("FRAMEDMARKER", PdfCore.PdfReadDocument.Open(output.ToArray()).ExtractText(), StringComparison.Ordinal);
        }
        RtfDocument plain = RtfDocument.Create();
        plain.AddParagraph("PLAINMARKER");
        Assert.DoesNotContain(plain.ToPdfDocumentResult(options).Warnings, warning => warning.Code == "ParagraphFrameFlattened");
        using WordDocument word = plain.ToWordDocumentResult().RequireNoLoss();
    }

    [Fact]
    public void SkippedAndInactiveRunningStoriesDoNotReportFrameFlattening() {
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("BODY");
        source.AddHeader(RtfHeaderFooterKind.FirstHeader).AddParagraph("INACTIVE").Frame.WidthTwips = 3600;
        source.AddHeader().AddParagraph("SKIPPED").Frame.WidthTwips = 3600;

        Assert.DoesNotContain(source.ToPdfDocumentResult(new RtfToPdfOptions { IncludeHeaderFooters = false }).Warnings,
            warning => warning.Code == "ParagraphFrameFlattened");
        Assert.Single(source.ToPdfDocumentResult().Warnings, warning => warning.Code == "ParagraphFrameFlattened");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordTabClearMetadataReportsOmissionEvenWithNoActiveTabStops(bool inStyle) {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("CLEARED");
        var tabs = new W.Tabs(new W.TabStop { Val = W.TabStopValues.Clear, Position = 4320 });
        if (inStyle) {
            word.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.StyleParagraphProperties(tabs)) {
                    Type = W.StyleValues.Paragraph, StyleId = "ClearedTabs"
                });
            paragraph.SetStyleId("ClearedTabs");
        } else {
            paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.AddChild(tabs, true);
        }
        string original = paragraph._paragraph.OuterXml;
        RtfConversionResult<RtfDocument> result = word.ToRtfDocumentResult();
        RtfConversionDiagnostic warning = Assert.Single(result.Report.Diagnostics,
            diagnostic => diagnostic.Code == "WordRtfTabClearsOmitted");
        Assert.Equal(RtfConversionAction.Omitted, warning.Action);
        Assert.Equal(1, warning.Count);
        Assert.True(result.HasLoss);
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
        Assert.Equal(original, paragraph._paragraph.OuterXml);
        Assert.Contains("CLEARED", string.Concat(result.Value.Paragraphs.Select(item => item.ToPlainText())), StringComparison.Ordinal);
        Assert.Empty(Assert.Single(result.Value.Paragraphs).TabStops);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void WordFramesReportFlatteningWhenTheirTextIsConvertedToRtf(bool inStyle) {
        using WordDocument word = WordDocument.Create();
        WordParagraph paragraph = word.AddParagraph("FRAMEDMARKER");
        var frame = new W.FrameProperties();
        frame.SetAttribute(new DocumentFormat.OpenXml.OpenXmlAttribute("w", "w",
            "http://schemas.openxmlformats.org/wordprocessingml/2006/main", "3600"));
        if (inStyle) {
            word.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(
                new W.Style(new W.StyleParagraphProperties(frame)) {
                    Type = W.StyleValues.Paragraph, StyleId = "Positioned"
                });
            paragraph.SetStyleId("Positioned");
        } else {
            paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.AddChild(frame, true);
        }
        string paragraphXml = paragraph._paragraph.OuterXml;
        string stylesXml = word.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml;

        RtfConversionResult<RtfDocument> result = word.ToRtfDocumentResult();

        RtfConversionDiagnostic diagnostic = Assert.Single(result.Report.Diagnostics,
            item => item.Code == "WordRtfParagraphFramesFlattened");
        Assert.Equal(RtfConversionAction.Flattened, diagnostic.Action);
        Assert.Equal(1, diagnostic.Count);
        Assert.True(result.HasLoss);
        Assert.Throws<RtfConversionLossException>(() => result.RequireNoLoss());
        Assert.Contains("FRAMEDMARKER", string.Concat(result.Value.Paragraphs.Select(item => item.ToPlainText())), StringComparison.Ordinal);
        Assert.Equal(paragraphXml, paragraph._paragraph.OuterXml);
        Assert.Equal(stylesXml, word.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.OuterXml);
    }

    private static RtfParagraph AddParagraph(RtfDocument document, string story) {
        switch (story) {
            case "body": return document.AddParagraph();
            case "table": return document.AddTable(1, 1).Rows[0].Cells[0].AddParagraph();
            case "header": return document.AddHeader().AddParagraph();
            case "footer": return document.AddFooter().AddParagraph();
            default:
                RtfNote note = document.AddNote(story == "endnote" ? RtfNoteKind.Endnote : RtfNoteKind.Footnote);
                document.AddParagraph().AddNoteReference(note);
                return note.AddParagraph();
        }
    }
}
