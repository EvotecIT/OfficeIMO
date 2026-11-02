using System;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordSmartArtSharedPageTests {
    [Theory]
    [InlineData(WordSmartArtType.BasicProcess, "process")]
    [InlineData(WordSmartArtType.Cycle, "cycle")]
    [InlineData(WordSmartArtType.Hierarchy, "hierarchy")]
    [InlineData(WordSmartArtType.PictureOrgChart, "hierarchy")]
    [InlineData(WordSmartArtType.CustomSmartArt1, "persisted")]
    public void SmartArtTextAndGeometrySurviveSharedPdfAndImagePages(WordSmartArtType type, string expectedProjection) {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("Before diagram");
        WordSmartArt smartArt = document.AddSmartArt(type);
        if (smartArt.CanModifyNodes)
            while (smartArt.NodeCount < (type == WordSmartArtType.Cycle ? 4 : 3)) smartArt.AddNode("Node");
        string[] labels = SetNodeLabels(smartArt);
        document.AddParagraph("After diagram");

        OfficeDrawing scene = Assert.IsType<OfficeDrawing>(WordSmartArtDrawingRenderer.CreateDrawing(smartArt, out string projection));
        Assert.Equal(expectedProjection, projection);
        var nodes = scene.Elements.OfType<OfficeDrawingText>().ToDictionary(text => text.Text);
        Assert.All(labels, label => Assert.True(nodes.ContainsKey(label), "Missing node " + label));
        if (projection == "process") {
            Assert.True(nodes[labels[0]].X < nodes[labels[1]].X);
            Assert.True(nodes[labels[1]].X < nodes[labels[2]].X);
            Assert.Equal(nodes[labels[0]].Y, nodes[labels[1]].Y, 5);
            Assert.Equal(labels.Length - 1, scene.Shapes.Count(shape => shape.Shape.StrokeEndMarker != null));
        } else if (projection == "cycle") {
            Assert.True(nodes[labels[0]].Y < nodes[labels[2]].Y);
            Assert.True(nodes[labels[1]].X > nodes[labels[3]].X);
            Assert.Equal(labels.Length, scene.Shapes.Count(shape => shape.Shape.StrokeEndMarker != null));
        } else if (projection == "hierarchy") {
            Assert.True(nodes[labels[0]].Y < nodes[labels[1]].Y);
            Assert.True(nodes[labels[1]].X < nodes[labels[2]].X);
        } else {
            Assert.True(nodes[labels[0]].X < nodes[labels[1]].X);
            Assert.True(nodes[labels[3]].Y > nodes[labels[0]].Y);
        }

        var conversion = document.ConvertToPdfWithEvidence();
        string warningCode = projection == "persisted" ? "native-smart-art-persisted-layout" : "native-smart-art-layout-approximated";
        Assert.Contains(conversion.Warnings, warning => warning.Code == warningCode &&
            warning.LossKind == OfficeConversionLossKind.Approximation);
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "native-body-smart-art-unsupported");
        AssertSharedSvgPages(document, conversion, labels.Concat(new[] { "Before diagram", "After diagram" }).ToArray(), warningCode);
    }

    [Fact]
    public void MappedSmartArtFallbackIsAllowedByStrictOmissionPolicy() {
        using WordDocument document = WordDocument.Create();
        WordSmartArt smartArt = document.AddSmartArt(WordSmartArtType.BasicProcess);
        while (smartArt.NodeCount < 3) smartArt.AddNode("Node");
        string[] labels = SetNodeLabels(smartArt);
        OfficeImageExportResult image = document.ExportImage(OfficeImageExportFormat.Svg,
            new WordImageExportOptions {
                Policy = new OfficeImageExportPolicy { RequireNoOmissions = true }
            });
        Assert.Contains(image.Diagnostics, diagnostic => diagnostic.Code == "native-smart-art-layout-approximated" &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
        Assert.DoesNotContain(image.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
        string svg = Encoding.UTF8.GetString(image.Bytes);
        Assert.All(labels, label => Assert.Contains(label, svg, StringComparison.Ordinal));
    }

    [Fact]
    public void UnknownDiagramLayoutRetainsAllNodeTextInAnExplicitFallback() {
        using WordDocument document = WordDocument.Create();
        WordSmartArt smartArt = document.AddSmartArt(WordSmartArtType.BasicProcess);
        while (smartArt.NodeCount < 3) smartArt.AddNode("Node");
        string[] labels = SetNodeLabels(smartArt);
        document.MainDocumentPartRoot.DiagramLayoutDefinitionParts.Single().LayoutDefinition!.UniqueId = "urn:officeimo:unknown-layout";
        var conversion = document.ConvertToPdfWithEvidence();
        Assert.Contains(conversion.Warnings, warning => warning.Code == "native-smart-art-layout-approximated" &&
            warning.Message.Contains("text layout", StringComparison.Ordinal));
        AssertSharedSvgPages(document, conversion, labels, "native-smart-art-layout-approximated");
    }

    [Fact]
    public void IncompletePersistedLayoutCannotDropAnEditableNode() {
        using WordDocument document = WordDocument.Create();
        WordSmartArt smartArt = document.AddSmartArt(WordSmartArtType.CustomSmartArt1);
        string[] labels = SetNodeLabels(smartArt);
        var part = document.MainDocumentPartRoot.DiagramPersistLayoutParts.Single();
        string modelId = smartArt.GetPersistedLayoutShapes().First(shape => shape.Text == labels[0]).ModelId;
        part.Drawing!.Descendants<DocumentFormat.OpenXml.Office.Drawing.Shape>()
            .First(shape => shape.ModelId?.Value == modelId).Remove();
        var conversion = document.ConvertToPdfWithEvidence();
        Assert.Contains(conversion.Warnings, warning => warning.Code == "native-smart-art-layout-approximated");
        AssertSharedSvgPages(document, conversion, labels, "native-smart-art-layout-approximated");
    }

    [Fact]
    public void SmartArtUsesNativePageFlowAndConservesFollowingText() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph(string.Join(" ", Enumerable.Repeat("Text before diagram.", 170)));
        WordSmartArt smartArt = document.AddSmartArt(WordSmartArtType.BasicProcess);
        while (smartArt.NodeCount < 3) smartArt.AddNode("Node");
        string[] labels = SetNodeLabels(smartArt);
        document.AddParagraph("AFTER_DIAGRAM");
        var conversion = document.ConvertToPdfWithEvidence();
        var read = PdfDocument.Load(new System.IO.MemoryStream(conversion.ToBytes()));
        Assert.True(read.Pages.Count > 1);
        int[] diagramPages = read.Pages.Select((page, index) => new { index,
            Text = string.Concat(page.GetTextSpans().Select(span => span.Text)) })
            .Where(page => labels.Any(label => page.Text.Contains(label, StringComparison.Ordinal)))
            .Select(page => page.index).ToArray();
        int diagramPage = Assert.Single(diagramPages);
        string diagramText = string.Concat(read.Pages[diagramPage].GetTextSpans().Select(span => span.Text));
        Assert.All(labels, label => Assert.Contains(label, diagramText, StringComparison.Ordinal));
        Assert.Contains("AFTER_DIAGRAM", string.Concat(read.Pages.Skip(diagramPage)
            .SelectMany(page => page.GetTextSpans()).Select(span => span.Text)), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(false, WordSmartArtType.BasicProcess, "native-smart-art-layout-approximated")]
    [InlineData(true, WordSmartArtType.BasicProcess, "native-smart-art-layout-approximated")]
    [InlineData(false, WordSmartArtType.CustomSmartArt1, "native-smart-art-persisted-layout")]
    [InlineData(true, WordSmartArtType.CustomSmartArt1, "native-smart-art-persisted-layout")]
    [InlineData(false, WordSmartArtType.CustomSmartArt2, "native-smart-art-persisted-layout")]
    [InlineData(true, WordSmartArtType.CustomSmartArt2, "native-smart-art-persisted-layout")]
    public void HeaderAndFooterSmartArtEnterSharedRunningContent(bool footer, WordSmartArtType type, string warningCode) {
        using WordDocument document = WordDocument.Create();
        WordHeaderFooter story = footer ? document.FooterDefaultOrCreate : document.HeaderDefaultOrCreate;
        WordSmartArt smartArt = story.AddParagraph().AddSmartArt(type);
        if (type == WordSmartArtType.BasicProcess)
            while (smartArt.NodeCount < 3) smartArt.AddNode("Node");
        string[] labels = SetNodeLabels(smartArt);
        DW.Extent extent = smartArt._drawing.Inline!.Extent!;
        extent.Cx = 2_743_200L;
        extent.Cy = 914_400L;
        document.AddParagraph("Body page one");
        document.AddPageBreak();
        document.AddParagraph("Body page two");
        using WordDocument reopened = WordDocument.Load(new System.IO.MemoryStream(document.ToBytes()));
        var conversion = reopened.ConvertToPdfWithEvidence();
        Assert.DoesNotContain(conversion.Warnings, warning => warning.Code == "native-header-footer-smart-art-unsupported");
        Assert.Contains(conversion.Warnings, warning => warning.Code == warningCode);
        var read = PdfReadDocument.Open(conversion.ToBytes());
        Assert.Equal(2, read.Pages.Count);
        Assert.All(read.Pages, page => {
            string text = string.Concat(page.GetTextSpans(includeArtifactText: true).Select(span => span.Text));
            Assert.All(labels, label => Assert.Contains(label, text, StringComparison.Ordinal));
        });
        Assert.All(reopened.CreateVisualSnapshots(), page => {
            string text = string.Concat(page.TextSpans.Select(span => span.Text));
            Assert.All(labels, label => Assert.Contains(label, text, StringComparison.Ordinal));
        });
        foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Png, OfficeImageExportFormat.Svg }) {
            var expected = conversion.ExportImages(format, new PdfImageExportOptions { Scale = .5D });
            var actual = reopened.ExportImages(format, new WordImageExportOptions { Scale = .5D });
            Assert.Equal(expected.Count, actual.Count);
            for (int index = 0; index < expected.Count; index++) Assert.Equal(expected[index].Bytes, actual[index].Bytes);
        }
    }

    private static string[] SetNodeLabels(WordSmartArt smartArt) {
        string[] labels = Enumerable.Range(0, smartArt.NodeCount).Select(index => "Node" + index).ToArray();
        smartArt.ReplaceTexts(labels);
        return labels;
    }

    private static void AssertSharedSvgPages(WordDocument document, PdfDocumentConversionResult conversion,
        string[] labels, string warningCode) {
        var expected = conversion.ExportImages(OfficeImageExportFormat.Svg, new PdfImageExportOptions { Scale = .5D });
        var actual = document.ExportImages(OfficeImageExportFormat.Svg, new WordImageExportOptions { Scale = .5D });
        Assert.Equal(expected.Count, actual.Count);
        for (int index = 0; index < expected.Count; index++) Assert.Equal(expected[index].Bytes, actual[index].Bytes);
        // SVG may paint glyph outlines or split runs at spaces. Its bytes must
        // match the PDF route; text recall comes from the actual page spans.
        string text = string.Concat(document.CreateVisualSnapshots().SelectMany(page => page.TextSpans).Select(span => span.Text));
        Assert.All(labels, label => Assert.Contains(label, text, StringComparison.Ordinal));
        Assert.Contains(actual.SelectMany(image => image.Diagnostics), diagnostic => diagnostic.Code == warningCode &&
            diagnostic.LossKind == OfficeConversionLossKind.Approximation);
    }
}
