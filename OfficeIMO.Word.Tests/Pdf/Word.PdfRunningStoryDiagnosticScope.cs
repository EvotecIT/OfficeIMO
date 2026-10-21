using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using V = DocumentFormat.OpenXml.Vml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_InactiveStoredRunningControlDoesNotReportAnOmission(bool evenHeader) {
        using WordDocument source = CreateRunningDiagnosticScopeDocument(evenHeader, active: false);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        Assert.Single(document._wordprocessingDocument.MainDocumentPart!.HeaderParts
            .SelectMany(part => part.Header.Descendants<W.SdtRun>()));
        var result = document.ToPdfDocumentResult(RunningDiagnosticScopeOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("ACTIVE", page.Text);
        Assert.Contains("BODY", page.Text);
        Assert.DoesNotContain("PREFIX", page.Text);
        Assert.Empty(page.GetImages());
        Assert.DoesNotContain(result.Warnings,
            warning => warning.Code == "NativeHeaderFooterContentControlUnsupported");
    }

    [Fact]
    public void SaveAsPdf_ActiveUnsupportedRunningControlReportsItsParagraphOnce() {
        using WordDocument source = CreateRunningDiagnosticScopeDocument(evenHeader: false, active: true);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        Assert.Empty(document.ValidateDocument());
        var result = document.ToPdfDocumentResult(RunningDiagnosticScopeOptions());
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        var page = Assert.Single(pdf.GetPages());
        Assert.Contains("PREFIX", page.Text);
        Assert.Contains("BODY", page.Text);
        var warning = Assert.Single(result.Warnings,
            item => item.Code == "NativeHeaderFooterContentControlUnsupported");
        Assert.Equal("first header", warning.Source);
        Assert.Equal(OfficeIMO.OfficeConversionLossKind.Omission, warning.LossKind);
        Assert.Throws<InvalidOperationException>(() => result.RequireNoLoss());
    }

    private static WordDocument CreateRunningDiagnosticScopeDocument(bool evenHeader, bool active) {
        WordDocument document = WordDocument.Create();
        document.DifferentFirstPage = true;
        document.DifferentOddAndEvenPages = true;
        document.AddHeadersAndFooters();
        RequireSectionHeader(document, 0, W.HeaderFooterValues.Default).AddParagraph("ACTIVE");
        WordHeader target = RequireSectionHeader(document, 0,
            evenHeader ? W.HeaderFooterValues.Even : W.HeaderFooterValues.First);
        target._header!.RemoveAllChildren();

        HeaderPart part = document._wordprocessingDocument.MainDocumentPart!.HeaderParts
            .Single(candidate => ReferenceEquals(candidate.Header, target._header));
        ImagePart image = part.AddImagePart(ImagePartType.Png);
        byte[] png = OfficeRasterImageEncoder.Encode(new OfficeRasterImage(4, 4, OfficeColor.Black), OfficeImageExportFormat.Png);
        using (var stream = new MemoryStream(png)) image.FeedData(stream);
        string relationshipId = part.GetIdOfPart(image);
        var picture = new W.Picture(new V.Shape(new V.ImageData { RelationshipId = relationshipId }) {
            Id = "SyntheticStoredPicture", Style = "width:12pt;height:12pt"
        });
        // This unsupported gallery carrier contains a valid, task-owned VML
        // image cache. It is distinct from supported w:picture SDT controls.
        var control = new W.SdtRun(new W.SdtProperties(new W.SdtId { Val = 321 },
            new W.SdtContentDocPartObject(new W.DocPartGallery { Val = "Synthetic gallery" })),
            new W.SdtContentRun(new W.Run(picture)));
        var paragraph = new W.Paragraph();
        for (int index = 0; index < 6; index++) {
            paragraph.Append(new W.Run(new W.Text("PREFIX ") {
                Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve
            }));
        }
        paragraph.Append(control);
        target._header.Append(paragraph);
        document.AddParagraph("BODY");
        // Keep the native parts and references intact while changing which
        // stored variants are eligible for this conversion.
        if (!active || evenHeader) document.Sections[0]._sectionProperties.RemoveAllChildren<W.TitlePage>();
        if (!active || !evenHeader) document._wordprocessingDocument.MainDocumentPart.DocumentSettingsPart!
            .Settings!.RemoveAllChildren<W.EvenAndOddHeaders>();
        return document;
    }

    private static WordToPdfOptions RunningDiagnosticScopeOptions() => new() {
        IncludePageNumbers = false,
        ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
    };
}
