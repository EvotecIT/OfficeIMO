using System;
using System.IO;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public sealed class PowerPointPdfSlideLinkTests {
    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void StrokedLineLinksHaveUsableRegionsForZeroAndOneEmuExtents(bool vertical, bool helper) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(240));
        PowerPointSlide first = presentation.AddSlide(), second = presentation.AddSlide();
        PowerPointAutoShape line = helper
            ? first.AddLinePoints(20, 40, vertical ? 20 : 220, vertical ? 200 : 40)
            : first.AddShapePoints(OfficePresetShapeType.Line, 20, 40, vertical ? 0 : 200, vertical ? 160 : 0);
        line.OutlineColor = "000000"; line.OutlineWidthPoints = 2;
        line.SetHyperlink(second);
        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        PdfCore.PdfDocumentConversionResult conversion = reopened.ConvertToPdfWithEvidence();
        PdfCore.PdfLinkAnnotation link = Assert.Single(PdfCore.PdfInspector.Inspect(conversion.ToBytes()).LinkAnnotations);
        Assert.Equal("powerpoint-slide-2", link.DestinationName);
        Assert.True(link.X2 - link.X1 >= 1.99D && link.Y2 - link.Y1 >= 1.99D);
        Assert.DoesNotContain(conversion.Report.FidelityDiagnostics, d => d.LossKind == OfficeConversionLossKind.Omission);
    }

    [Theory]
    [InlineData(30)]
    [InlineData(90)]
    public void RotatedLineLinksKeepTheStrokeCenterAndPageClip(double rotation) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(240));
        PowerPointSlide first = presentation.AddSlide(), second = presentation.AddSlide();
        PowerPointAutoShape line = first.AddShapePoints(OfficePresetShapeType.StraightConnector1, 20, 40, 200, 0);
        line.OutlineColor = "000000"; line.OutlineWidthPoints = 2;
        line.Rotation = rotation; line.SetHyperlink(second);
        PdfCore.PdfLinkAnnotation link = Assert.Single(PdfCore.PdfInspector.Inspect(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })).LinkAnnotations);
        OfficeTransform transform = new OfficeImageFrameTransform(rotation, 120, 40, false, false).CreateDestinationTransform();
        OfficePoint[] corners = new[] { new OfficePoint(19, 39), new OfficePoint(221, 39),
            new OfficePoint(19, 41), new OfficePoint(221, 41) }.Select(transform.TransformPoint).ToArray();
        Assert.Equal("powerpoint-slide-2", link.DestinationName);
        Assert.InRange(Math.Abs(Math.Max(0, corners.Min(p => p.X)) - link.X1), 0D, 0.002D);
        Assert.InRange(Math.Abs(Math.Min(320, corners.Max(p => p.X)) - link.X2), 0D, 0.002D);
        Assert.InRange(Math.Abs(240 - Math.Min(240, corners.Max(p => p.Y)) - link.Y1), 0D, 0.002D);
        Assert.InRange(Math.Abs(240 - Math.Max(0, corners.Min(p => p.Y)) - link.Y2), 0D, 0.002D);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void InheritedNextSlideRunsUseEachActiveSlideAndRetainTheRelationshipOwner(bool master) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide first = presentation.AddSlide(); presentation.AddSlide();
        PowerPointTextBox box = first.AddTextBoxPoints("Inherited next", 30, 30, 240, 90);
        box.PlaceholderType = null;
        A.Run run = box.Element.Descendants<A.Run>().Single();
        run.RunProperties ??= new A.RunProperties();
        run.RunProperties.Append(new A.HyperlinkOnClick {
            Id = string.Empty, Action = "ppaction://hlinkshowjump?jump=nextslide"
        });
        P.Shape native = (P.Shape)box.Element;
        native.Remove();
        var layout = first.SlidePart.SlideLayoutPart!;
        (master ? layout.SlideMasterPart!.SlideMaster!.CommonSlideData!.ShapeTree!
            : layout.SlideLayout!.CommonSlideData!.ShapeTree!).Append(native);
        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        for (int index = 0; index < 2; index++) {
            PowerPointTextRun inherited = reopened.Slides[index].GetInheritedShapesForExport()
                .OfType<PowerPointTextBox>().Single(b => b.Text == "Inherited next").Paragraphs[0].Runs[0];
            Assert.Equal("#slide-" + (index + 2), inherited.Hyperlink?.OriginalString);
            Assert.True(master ? inherited.OwnerPart is DocumentFormat.OpenXml.Packaging.SlideMasterPart
                : inherited.OwnerPart is DocumentFormat.OpenXml.Packaging.SlideLayoutPart);
        }
        PdfCore.PdfDocumentConversionResult conversion = reopened.ConvertToPdfWithEvidence();
        PdfCore.PdfLinkAnnotation link = Assert.Single(PdfCore.PdfInspector.Inspect(conversion.ToBytes()).LinkAnnotations);
        Assert.Equal("powerpoint-slide-2", link.DestinationName);
        Assert.Single(conversion.Report.FidelityDiagnostics, d => d.Code == "internal-slide-target-omitted"
            && d.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void ApplePptxExportKeepsUuidResolvedAndNextActionLinkTargetsInPdf() {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "IWorkCorpus", "native-exports",
            "keynote-slide-links-v15.4.pptx");
        using var presentation = PowerPointDocument.Load(path);
        PdfCore.PdfDocumentInfo pdf = PdfCore.PdfInspector.Inspect(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(3, pdf.Pages.Count);
        Assert.Equal(new int?[] { 2, 2, 1 }, pdf.LinkAnnotations.Select(link => pdf.NamedDestinations.Single(
            destination => destination.Name == link.DestinationName).PageNumber));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NativeNextSlideActionsWithoutRelationshipsResolveOrReportAMissingTarget(bool finalSlide) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide first = presentation.AddSlide(), second = presentation.AddSlide();
        PowerPointTextBox box = (finalSlide ? second : first).AddTextBoxPoints("Next slide", 30, 30, 240, 90);
        ((P.Shape)box.Element).NonVisualShapeProperties!.NonVisualDrawingProperties!.Append(
            new A.HyperlinkOnClick { Id = string.Empty, Action = "ppaction://hlinkshowjump?jump=nextslide" });
        A.Run run = box.Element.Descendants<A.Run>().Single();
        run.RunProperties ??= new A.RunProperties();
        run.RunProperties.Append(new A.HyperlinkOnClick {
            Id = string.Empty, Action = "ppaction://hlinkshowjump?jump=nextslide"
        });
        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        Assert.Equal(finalSlide ? "#slide-3" : "#slide-2",
            reopened.Slides[finalSlide ? 1 : 0].TextBoxes.Single().Hyperlink!.OriginalString);
        PdfCore.PdfDocumentConversionResult conversion = reopened.ConvertToPdfWithEvidence();
        PdfCore.PdfDocumentInfo pdf = PdfCore.PdfInspector.Inspect(conversion.ToBytes());
        if (finalSlide) {
            Assert.Empty(pdf.LinkAnnotations);
            Assert.Equal(2, conversion.Report.FidelityDiagnostics.Count(d => d.Code == "internal-slide-target-omitted"
                && d.LossKind == OfficeConversionLossKind.Omission));
        } else {
            Assert.Equal(2, pdf.LinkAnnotations.Count);
            Assert.All(pdf.LinkAnnotations, link => Assert.Equal("powerpoint-slide-2", link.DestinationName));
            Assert.DoesNotContain(conversion.Report.FidelityDiagnostics, d => d.LossKind == OfficeConversionLossKind.Omission);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void HiddenSlidesKeepOriginalTargetIdentityAndOmittedTargetsRemainDiagnosed(bool includeHidden) {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        PowerPointSlide first = presentation.AddSlide();
        PowerPointSlide hidden = presentation.AddSlide(); hidden.Hidden = true;
        PowerPointSlide visible = presentation.AddSlide();
        PowerPointTextBox shown = first.AddTextBoxPoints("Visible target", 20, 20, 300, 60);
        shown.Paragraphs[0].Runs[0].SetHyperlink(visible);
        PowerPointTextBox omitted = first.AddTextBoxPoints("Hidden target", 20, 100, 300, 60);
        omitted.Paragraphs[0].Runs[0].SetHyperlink(hidden);
        omitted.SetHyperlink(hidden);

        var options = new PowerPointPdfExportOptions { IncludeHiddenSlides = includeHidden };
        PdfCore.PdfDocumentConversionResult conversion = presentation.ConvertToPdfWithEvidence(options);
        PdfCore.PdfDocumentInfo pdf = PdfCore.PdfInspector.Inspect(conversion.ToBytes());
        Assert.Equal(includeHidden ? 3 : 2, pdf.Pages.Count);
        Assert.Equal(includeHidden ? 3 : 2, Assert.Single(pdf.NamedDestinations,
            d => d.Name == "powerpoint-slide-3").PageNumber);
        Assert.Contains(pdf.LinkAnnotations, link => link.DestinationName == "powerpoint-slide-3");
        if (includeHidden) {
            Assert.Equal(2, pdf.LinkAnnotations.Count(link => link.DestinationName == "powerpoint-slide-2"));
            Assert.DoesNotContain(conversion.Report.FidelityDiagnostics, d => d.Code == "internal-slide-target-omitted");
        } else {
            Assert.DoesNotContain(pdf.LinkAnnotations, link => link.DestinationName == "powerpoint-slide-2");
            Assert.Equal(2, conversion.Report.FidelityDiagnostics.Count(d => d.Code == "internal-slide-target-omitted"
                && d.LossKind == OfficeConversionLossKind.Omission));
        }
        // Reusing caller options must not retain an earlier operation's destination set.
        options.IncludeHiddenSlides = !includeHidden;
        PdfCore.PdfDocumentInfo opposite = PdfCore.PdfInspector.Inspect(presentation.ToPdf(options).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(!includeHidden, opposite.NamedDestinations.Any(d => d.Name == "powerpoint-slide-2"));
    }

    [Fact]
    public void AdjacentIdenticallyStyledRunsKeepDifferentDestinationsAfterSaveAndWrapping() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(360),OfficeIMO.OfficeLength.Points(240));
        PowerPointSlide first = presentation.AddSlide();
        PowerPointSlide second = presentation.AddSlide();
        PowerPointSlide third = presentation.AddSlide();
        PowerPointTextBox box = first.AddTextBoxPoints("FIRST ", 30, 30, 180, 160);
        box.FontName = "Arial"; box.FontSize = 16;
        box.Paragraphs[0].Runs[0].SetHyperlink(second);
        box.Paragraphs[0].AddRun("SECOND THIRD FOURTH FIFTH SIXTH").SetHyperlink(third);
        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        PdfCore.PdfDocumentConversionResult conversion = reopened.ConvertToPdfWithEvidence();
        byte[] bytes = conversion.ToBytes();
        PdfCore.PdfDocumentInfo pdf = PdfCore.PdfInspector.Inspect(bytes);
        PdfCore.PdfLinkAnnotation firstLink = Assert.Single(pdf.LinkAnnotations, l => l.DestinationName == "powerpoint-slide-2");
        var later = pdf.LinkAnnotations.Where(l => l.DestinationName == "powerpoint-slide-3").ToArray();
        Assert.True(later.Length >= 2, "The linked text must wrap to multiple clickable lines.");
        Assert.All(later, link => Assert.InRange(link.X1, 30D, 210D));
        Assert.True(firstLink.X2 <= later[0].X1 + 0.01D);
        Assert.DoesNotContain(conversion.Report.FidelityDiagnostics, d => d.LossKind == OfficeConversionLossKind.Omission);
        using var rendered = UglyToad.PdfPig.PdfDocument.Open(bytes);
        Assert.Contains("FIRST", rendered.GetPage(1).Text);
        Assert.Contains("SIXTH", rendered.GetPage(1).Text);
    }

    [Fact]
    public void PictureAndTableRunSlideLinksSurvivePptxSaveAndPdfReopen() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(400),OfficeIMO.OfficeLength.Points(240));
        PowerPointSlide first = presentation.AddSlide();
        PowerPointSlide second = presentation.AddSlide();
        PowerPointPicture picture = first.AddImagePoints(new MemoryStream(PdfPngTestImages.Create16BitRgbPng()),
            OfficeImageFormat.Png, 20, 20, 60, 40);
        picture.Rotation = 30; picture.SetHyperlink(second);
        PowerPointTableCell cell = first.AddTable(1,1,global::OfficeIMO.OfficeLength.Points(20),global::OfficeIMO.OfficeLength.Points(100),global::OfficeIMO.OfficeLength.Points(220),global::OfficeIMO.OfficeLength.Points(90)).GetCell(0, 0);
        cell.Text = "Cell target"; cell.Paragraphs[0].Runs[0].SetHyperlink(second);
        using var saved = new MemoryStream(); presentation.Save(saved); saved.Position = 0;
        using var reopened = PowerPointDocument.Load(saved);
        PdfCore.PdfDocumentConversionResult conversion = reopened.ConvertToPdfWithEvidence();
        PdfCore.PdfDocumentInfo pdf = PdfCore.PdfInspector.Inspect(conversion.ToBytes());
        Assert.Equal(2, pdf.LinkAnnotations.Count);
        Assert.All(pdf.LinkAnnotations, link => Assert.Equal("powerpoint-slide-2", link.DestinationName));
        PdfCore.PdfLinkAnnotation frame = pdf.LinkAnnotations.OrderByDescending(link => link.Y2).First();
        OfficeTransform transform = new OfficeImageFrameTransform(30, 50, 40).CreateDestinationTransform();
        OfficePoint[] corners = new[] { new OfficePoint(20, 20), new OfficePoint(80, 20),
            new OfficePoint(80, 60), new OfficePoint(20, 60) }.Select(transform.TransformPoint).ToArray();
        Assert.InRange(Math.Abs(corners.Min(p => p.X) - frame.X1), 0D, 0.002D);
        Assert.InRange(Math.Abs(corners.Max(p => p.X) - frame.X2), 0D, 0.002D);
        Assert.InRange(Math.Abs(240 - corners.Max(p => p.Y) - frame.Y1), 0D, 0.002D);
        Assert.DoesNotContain(conversion.Report.FidelityDiagnostics, d => d.LossKind == OfficeConversionLossKind.Omission);
    }

    [Fact]
    public void GroupTransformsMoveInternalLinkRectanglesWithTheLinkedGlyphs() {
        using var presentation = PowerPointDocument.Create(new MemoryStream());
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(320),OfficeIMO.OfficeLength.Points(220));
        PowerPointSlide first = presentation.AddSlide(); PowerPointSlide second = presentation.AddSlide();
        PowerPointTextBox box = first.AddTextBoxPoints("GROUP LINK", 40, 40, 200, 60);
        box.FontName = "Arial"; box.FontSize = 14;
        box.Paragraphs[0].Runs[0].SetHyperlink(second);
        PowerPointGroupShape group = first.GroupShapes(new PowerPointShape[] { box, first.AddRectanglePoints(40, 110, 200, 10) });
        PdfCore.PdfLinkAnnotation original = Assert.Single(PdfCore.PdfInspector.Inspect(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })).LinkAnnotations);
        group.Rotation = 30; group.HorizontalFlip = true;
        PdfCore.PdfLinkAnnotation actual = Assert.Single(PdfCore.PdfInspector.Inspect(presentation.ToPdf().ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true })).LinkAnnotations);
        OfficeTransform matrix = new OfficeImageFrameTransform(30, 140, 80, true, false).CreateDestinationTransform();
        OfficePoint[] corners = new[] { new OfficePoint(original.X1, 220 - original.Y1),
            new OfficePoint(original.X1, 220 - original.Y2), new OfficePoint(original.X2, 220 - original.Y1),
            new OfficePoint(original.X2, 220 - original.Y2) }.Select(matrix.TransformPoint).ToArray();
        Assert.Equal("powerpoint-slide-2", actual.DestinationName);
        Assert.InRange(Math.Abs(corners.Min(p => p.X) - actual.X1), 0D, 0.002D);
        Assert.InRange(Math.Abs(corners.Max(p => p.X) - actual.X2), 0D, 0.002D);
        Assert.InRange(Math.Abs(220 - corners.Max(p => p.Y) - actual.Y1), 0D, 0.002D);
        Assert.InRange(Math.Abs(220 - corners.Min(p => p.Y) - actual.Y2), 0D, 0.002D);
    }
}
