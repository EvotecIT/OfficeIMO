using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using W = DocumentFormat.OpenXml.Wordprocessing;
using W14 = DocumentFormat.OpenXml.Office2010.Word.Drawing;
using Wps = DocumentFormat.OpenXml.Office2010.Word.DrawingShape;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_ForegroundRunningBoxMeasuresManyOrdinaryBodyRuns() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        W.Paragraph body = SetAnchoredControlBodyAutoSpacing(document);
        var paragraph = new WordParagraph(document, body);
        paragraph.LineSpacingPoints = 6;
        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        foreach (int index in Enumerable.Range(0, 300))
            body.Append(new W.Run(new W.Text("x")));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;

        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());

        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains("S2", string.Join(" ", Enumerable.Range(1, pdf.NumberOfPages).Select(index => pdf.GetPage(index).Text)));
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_ForegroundRunningBoxWithUpwardAutoBodyInkRetainsWarnedFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        W.Paragraph body = SetAnchoredControlBodyAutoSpacing(document);
        body.ParagraphProperties!.SpacingBetweenLines!.Line = "120";
        W.Run run = body.Elements<W.Run>().First();
        run.RunProperties ??= new W.RunProperties();
        run.RunProperties.FontSize = new W.FontSize { Val = "288" };
        DW.Anchor anchor = document.Sections[1].Header!.First!.TextBoxes.Single().Anchor!;
        anchor.Elements<W14.RelativeHeight>().Single().PercentageHeight!.Text = "8000";
        SetAnchoredControlCyanFill(anchor);
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        string[] storiesBefore = AnchoredRunningStoryXml(document);
        WordPdfExportOptions options = EmptyRunningOptions();
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        var result = document.ConvertToPdfWithEvidence(options);
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.All(AnchoredRunningLetters(pdf.GetPage(2), "S2"), letter => Assert.Equal(144D, letter.PointSize, 3));
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Equal(storiesBefore, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_ForegroundRunningBoxWithOrdinaryBodyHighlightRetainsWarnedFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        W.Paragraph body = SetAnchoredControlBodyAutoSpacing(document);
        W.Run run = body.Elements<W.Run>().First();
        run.RunProperties ??= new W.RunProperties();
        run.RunProperties.FontSize = new W.FontSize { Val = "24" };
        run.RunProperties.Highlight = new W.Highlight { Val = W.HighlightColorValues.Yellow };
        run.Elements<W.Text>().First().Text += " ORDINARY HIGHLIGHTED BODY PAINT";
        WordTextBox box = document.Sections[1].Header!.First!.TextBoxes.Single();
        box.Anchor!.RemoveAllChildren<W14.RelativeHeight>();
        box.Height = 571500L;
        box.Anchor.Descendants<A.Transform2D>().Single().Extents!.Cy = 571500L;
        box.Anchor.VerticalPosition!.PositionOffset!.Text = "341630";
        foreach (W.Break lineBreak in box.Paragraphs.Single()._paragraph.Descendants<W.Break>().ToArray()) lineBreak.Remove();
        SetAnchoredControlCyanFill(box.Anchor);
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        string[] storiesBefore = AnchoredRunningStoryXml(document);
        WordPdfExportOptions options = EmptyRunningOptions();
        options.ResourcePolicy.AllowDocumentFontEmbedding = true;
        var result = document.ConvertToPdfWithEvidence(options);
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.All(AnchoredRunningLetters(pdf.GetPage(2), "S2"), letter => Assert.Equal(12D, letter.PointSize, 3));
        Assert.Contains(pdf.GetPage(2).Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 1D, 0D));
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
        Assert.Equal(storiesBefore, AnchoredRunningStoryXml(document));
    }

    private static W.Paragraph SetAnchoredControlBodyAutoSpacing(WordDocument document) {
        W.Body body = document._wordprocessingDocument.MainDocumentPart!.Document.Body!;
        foreach (W.Paragraph paragraph in body.Elements<W.Paragraph>()) {
            paragraph.ParagraphProperties ??= new W.ParagraphProperties();
            paragraph.ParagraphProperties.SpacingBetweenLines = new W.SpacingBetweenLines {
                Line = "240", LineRule = W.LineSpacingRuleValues.Auto, Before = "0", After = "0"
            };
        }
        return body.Descendants<W.Text>().Single(text => text.Text == "S2").Ancestors<W.Paragraph>().First();
    }

    private static void SetAnchoredControlCyanFill(DW.Anchor anchor) {
        Wps.ShapeProperties shape = anchor.Descendants<Wps.ShapeProperties>().Single();
        shape.RemoveAllChildren<A.NoFill>();
        shape.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = "00FFFF" }), true);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ForegroundRunningBoxWithUpwardExactBodyInkRetainsWarnedFlow(bool inheritedSpacing) {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        DW.Anchor anchor = document.Sections[1].Header!.First!.TextBoxes.Single().Anchor!;
        anchor.Elements<W14.RelativeHeight>().Single().PercentageHeight!.Text = "8000";
        Wps.ShapeProperties shape = anchor.Descendants<Wps.ShapeProperties>().Single();
        shape.RemoveAllChildren<A.NoFill>();
        shape.AddChild(new A.SolidFill(new A.RgbColorModelHex { Val = "FFFF00" }), true);
        var main = document._wordprocessingDocument.MainDocumentPart!;
        W.Paragraph bodySource = main.Document.Body!.Descendants<W.Text>().Single(text => text.Text == "S2")
            .Ancestors<W.Paragraph>().First();
        var body = new WordParagraph(document, bodySource);
        W.Run bodyRun = bodySource.Elements<W.Run>().First();
        bodyRun.RunProperties ??= new W.RunProperties();
        bodyRun.RunProperties.FontSize = new W.FontSize { Val = "288" };
        if (inheritedSpacing) {
            main.StyleDefinitionsPart!.Styles!.Append(new W.Style {
                Type = W.StyleValues.Paragraph, StyleId = "UpwardBody",
                StyleParagraphProperties = new W.StyleParagraphProperties(
                    new W.SpacingBetweenLines { Line = "120", LineRule = W.LineSpacingRuleValues.Exact })
            });
            bodySource.ParagraphProperties!.RemoveAllChildren<W.SpacingBetweenLines>();
            bodySource.ParagraphProperties.ParagraphStyleId = new W.ParagraphStyleId { Val = "UpwardBody" };
        } else {
            body.LineSpacingPoints = 6;
            body.LineSpacingRule = WordLineSpacingRule.Exact;
        }
        string[] before = AnchoredRunningStoryXml(document);
        string stylesBefore = main.StyleDefinitionsPart!.Styles!.OuterXml;
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.All(AnchoredRunningLetters(pdf.GetPage(2), "S2"), letter => Assert.Equal(144D, letter.PointSize, 3));
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
        Assert.Equal(stylesBefore, main.StyleDefinitionsPart.Styles.OuterXml);
    }

    [Fact]
    public void SaveAsPdf_ForegroundRunningBoxOverNegativeBodyIndentRetainsWarnedFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        WordTextBox box = document.Sections[1].Header!.First!.TextBoxes.Single();
        box.Anchor!.RemoveAllChildren<W14.RelativeWidth>();
        box.Anchor.RemoveAllChildren<W14.RelativeHeight>();
        box.Width = box.Height = 571500L;
        box.Anchor.Descendants<A.Transform2D>().Single().Extents!.Cx = 571500L;
        box.Anchor.Descendants<A.Transform2D>().Single().Extents!.Cy = 571500L;
        DW.HorizontalPosition position = box.Anchor.HorizontalPosition!;
        position.RemoveAllChildren<DW.HorizontalAlignment>();
        position.PositionOffset = new DW.PositionOffset("127000");
        box.Anchor.VerticalPosition!.PositionOffset!.Text = "914400";
        W.Text text = document._wordprocessingDocument.MainDocumentPart!.Document.Body!
            .Descendants<W.Text>().Single(value => value.Text == "S2");
        W.Paragraph bodySource = text.Ancestors<W.Paragraph>().First();
        bodySource.ParagraphProperties!.Indentation = new W.Indentation { Left = "-1440" };
        text.Text += " NEGATIVE INDENT OVERPAINT";
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        foreach (string token in new[] { "PREFIX", "BOX", "SUFFIX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }

    [Fact]
    public void SaveAsPdf_DecorationOnlyNbspRunningBoxOverBodyRetainsWarnedFlow() {
        using WordDocument document = CreateAnchoredRunningTextBoxControl(includeTextBox: true);
        WordTextBox box = document.Sections[1].Header!.First!.TextBoxes.Single();
        box.Anchor!.VerticalPosition!.PositionOffset!.Text = "914400";
        WordParagraph inner = box.Paragraphs.Single();
        foreach (W.Break lineBreak in inner._paragraph.Descendants<W.Break>().ToArray()) lineBreak.Remove();
        inner.Text = new string('\u00a0', 10);
        W.Run source = inner._paragraph.Elements<W.Run>().First();
        source.RunProperties ??= new W.RunProperties();
        source.RunProperties.Highlight = new W.Highlight { Val = W.HighlightColorValues.Yellow };
        string[] before = AnchoredRunningStoryXml(document);
        var result = document.ConvertToPdfWithEvidence(EmptyRunningOptions());
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "native-mixed-text-box-layout-approximated");
        using var pdf = PdfPigDocument.Open(result.Value.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.NumberOfPages);
        foreach (string token in new[] { "PREFIX", "SUFFIX", "S2" }) AnchoredRunningLetters(pdf.GetPage(2), token);
        Assert.Contains(pdf.GetPage(2).Paths, path => path.IsFilled && path.FillColor.ToRGBValues() == (1D, 1D, 0D));
        Assert.Contains("COUNT 2", pdf.GetPage(3).Text);
        Assert.Equal(before, AnchoredRunningStoryXml(document));
    }
}
