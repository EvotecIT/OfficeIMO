using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.Tests.Pdf;
using Xunit;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Html.Tests;

public sealed class HtmlEditableLayoutPowerPointTests {
    [Theory]
    [InlineData("section", false)]
    [InlineData("article", true)]
    public void RootGroupingRegionKeepsItsOwningSemanticSlide(string groupingElement, bool wrapInMain) {
        string second = "<" + groupingElement + " style='position:absolute;width:180px;height:50px'>"
            + "<span>Second slide region</span></" + groupingElement + ">";
        if (wrapInMain) second = "<main>" + second + "</main>";
        const string first = "<section><p>First slide</p></section>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(first + second)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.Equal(2, presentation.Slides.Count);
        Assert.Contains(presentation.Slides[0].TextBoxes, box => box.Text == "First slide");
        Assert.Contains(presentation.Slides[1].TextBoxes, box => box.Text == "Second slide region");
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Message.Contains("owning semantic slide was not created", StringComparison.Ordinal));
    }

    [Fact]
    public void NegativeZIndexRegionIsPlacedBehindSemanticSlideContent() {
        const string html = "<div style='position:absolute;z-index:-2;width:140px;height:60px'>Far behind</div>" +
            "<div style='position:absolute;z-index:-1;width:140px;height:60px'>Near behind</div>" +
            "<p>Front semantic content</p>";
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);
        PowerPointTextBox farBehind = Assert.Single(slide.TextBoxes, box => box.Text == "Far behind");
        PowerPointTextBox nearBehind = Assert.Single(slide.TextBoxes, box => box.Text == "Near behind");
        PowerPointTextBox front = Assert.Single(slide.TextBoxes, box => box.Text == "Front semantic content");

        Assert.True(farBehind.DrawingOrder < nearBehind.DrawingOrder);
        Assert.True(nearBehind.DrawingOrder < front.DrawingOrder);
    }

    [Fact]
    public void MixedFlowStackingSimplificationIsDiagnosed() {
        const string html = "<div style='position:absolute;width:140px;height:60px'>Early region</div>" +
            "<p>Later semantic content</p>";
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Detail != null
            && diagnostic.Detail.Contains("stacking=appended-after-semantic-content", StringComparison.Ordinal));
    }

    [Fact]
    public void OpaqueRegionFillIsPaintedBeforeItsPictures() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:140px;height:60px;background:#ff0000'>Region" +
            "<img src='" + image + "' style='width:12px;height:12px'></div>";
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using PresentationDocument package = PresentationDocument.Open(new MemoryStream(stream.ToArray()), false);
        P.ShapeTree tree = package.PresentationPart!.SlideParts.Single().Slide.CommonSlideData!.ShapeTree!;
        int regionIndex = tree.ChildElements
            .Select((element, index) => (element, index))
            .Single(item => item.element is P.Shape shape
                && shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value?.StartsWith("HTML ", StringComparison.Ordinal) == true)
            .index;
        int pictureIndex = tree.ChildElements
            .Select((element, index) => (element, index))
            .Single(item => item.element is P.Picture)
            .index;
        Assert.True(regionIndex < pictureIndex);
    }

    [Fact]
    public void RegionBackgroundPictureIsPaintedBehindItsTextBox() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<style>.background{position:absolute;width:140px;height:60px;background-image:url('" + image +
            "');background-repeat:no-repeat;background-size:140px 60px}</style>" +
            "<div class='background'>Text over background</div>";
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        PowerPointSlide generatedSlide = Assert.Single(result.Value.Slides);
        PowerPointPicture background = Assert.Single(generatedSlide.Pictures);
        PowerPointTextBox text = Assert.Single(generatedSlide.TextBoxes, box =>
            box.Name?.StartsWith("HTML ", StringComparison.Ordinal) == true);
        Assert.True(background.DrawingOrder < text.DrawingOrder);
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using PowerPointPresentation reopened = PowerPointPresentation.Load(
            new MemoryStream(stream.ToArray()),
            new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        Assert.Single(Assert.Single(reopened.Slides).Pictures);
    }

    [Fact]
    public void ProjectedPicturesHonorImportPicturesOption() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:140px;height:60px'>Region" +
            "<img src='" + image + "' style='width:12px;height:12px'></div>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions {
                Mode = HtmlImportMode.Generic,
                ImportPictures = false
            });
        using PowerPointPresentation presentation = result.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);

        Assert.Empty(slide.Pictures);
        Assert.Contains(slide.TextBoxes, box => box.Text == "Region");
    }

    [Fact]
    public void ProjectedRegionUsesItsOwningSemanticSlide() {
        const string html = "<section><h1>First slide</h1><p>First body</p></section>" +
            "<section><h1>Second slide</h1><div style='position:absolute;width:140px;height:60px'>Second region</div></section>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.Equal(2, presentation.Slides.Count);
        Assert.DoesNotContain(presentation.Slides[0].TextBoxes, box => box.Text == "Second region");
        Assert.Contains(presentation.Slides[1].TextBoxes, box => box.Text == "Second region");
    }

    [Fact]
    public void LaterSemanticSlideUsesSectionLocalProjectedCoordinates() {
        const string html = "<section style='height:900px'><h1>First slide</h1></section>" +
            "<section><h1>Second slide</h1><div style='position:absolute;left:20px;top:30px;width:140px;height:60px'>Local region</div></section>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.Equal(2, presentation.Slides.Count);
        PowerPointTextBox region = Assert.Single(presentation.Slides[1].TextBoxes, box => box.Text == "Local region");
        Assert.InRange(region.LeftPoints, 0D, 150D);
        Assert.InRange(region.TopPoints, 0D, 150D);
    }

    [Fact]
    public void ExplicitGenericModeProjectsRegionsInsideSemanticLookingSlideClasses() {
        const string html = "<section class='officeimo-slide'><h1>Generic slide</h1>" +
            "<div style='position:absolute;width:140px;height:60px'>Generic region</div></section>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.Contains(Assert.Single(presentation.Slides).TextBoxes, box => box.Text == "Generic region");
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
    }

    [Fact]
    public void ProjectedRegionHonorsMetadataAndGeometryLimits() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;left:900px;top:800px;width:700px;height:600px'>Short" +
            "<img src='" + image + "' style='width:500px;height:400px'></div>" +
            "<div style='position:absolute;width:120px;height:40px'>This metadata is much too long</div>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxMetadataCharacters = 8;
        limits.MaxAbsoluteGeometry = 50D;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using PowerPointPresentation presentation = result.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);

        Assert.All(slide.TextBoxes.Where(box => box.Text == "Short"), box => {
            Assert.InRange(box.LeftPoints, -50D, 50D);
            Assert.InRange(box.TopPoints, -50D, 50D);
            Assert.InRange(box.WidthPoints, 1D, 50D);
            Assert.InRange(box.HeightPoints, 1D, 50D);
        });
        Assert.DoesNotContain(slide.TextBoxes, box => box.Text.Contains("metadata is much", StringComparison.Ordinal));
        Assert.All(slide.Pictures, picture => {
            Assert.InRange(picture.LeftPoints, -50D, 50D);
            Assert.InRange(picture.TopPoints, -50D, 50D);
            Assert.InRange(picture.WidthPoints, 1D, 50D);
            Assert.InRange(picture.HeightPoints, 1D, 50D);
        });
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.SemanticValueInvalid);
    }

    [Fact]
    public void OmittedPositionedRegionDoesNotDisplaceFollowingInFlowRegion() {
        const string html = "<div style='position:absolute;left:0;top:0;width:300px;height:80px'>Metadata too long</div>" +
            "<div style='display:grid;width:300px;height:80px'>Grid</div>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxMetadataCharacters = 4;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        HtmlToPowerPointResult baseline = HtmlConversionDocument.Parse(
                "<div style='display:grid;width:300px;height:80px'>Grid</div>")
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using PowerPointPresentation presentation = result.Value;
        using PowerPointPresentation baselinePresentation = baseline.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);

        Assert.DoesNotContain(slide.TextBoxes, box => box.Text.Contains("Metadata", StringComparison.Ordinal));
        PowerPointTextBox grid = Assert.Single(slide.TextBoxes, box => box.Text == "Grid");
        PowerPointTextBox baselineGrid = Assert.Single(
            Assert.Single(baselinePresentation.Slides).TextBoxes, box => box.Text == "Grid");
        Assert.Equal(baselineGrid.TopPoints, grid.TopPoints, precision: 3);
        Assert.DoesNotContain(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified
            && diagnostic.Message.Contains("moved", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void ShapeBudgetOmissionDoesNotReservePhantomPositionedBounds() {
        const string gridOnly = "<div style='display:grid;width:300px;height:80px'>Grid</div>";
        const string html = gridOnly
            + "<div style='position:absolute;z-index:10;left:0;top:0;width:300px;height:80px'>Omitted</div>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxShapes = 2;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        HtmlToPowerPointResult baseline = HtmlConversionDocument.Parse(gridOnly)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using PowerPointPresentation presentation = result.Value;
        using PowerPointPresentation baselinePresentation = baseline.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);

        Assert.DoesNotContain(slide.TextBoxes, box => box.Text == "Omitted");
        PowerPointTextBox grid = Assert.Single(slide.TextBoxes, box => box.Text == "Grid");
        PowerPointTextBox baselineGrid = Assert.Single(
            Assert.Single(baselinePresentation.Slides).TextBoxes, box => box.Text == "Grid");
        Assert.Equal(baselineGrid.TopPoints, grid.TopPoints, precision: 3);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    [Fact]
    public void PositionedAndGridRegionsReopenAsEditableSlideGeometry() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(4, 3));
        string html = "<style>" +
            ".positioned{position:absolute;left:32px;top:24px;width:240px;height:72px;background-color:#dbeafe;" +
            "background-image:url('" + image + "'),url('" + image + "');background-repeat:no-repeat;background-size:18px 18px;" +
            "background-position:left top,right bottom;box-shadow:2px 2px 4px #555,4px 4px 8px #999}" +
            ".grid{display:grid;grid-template-columns:1fr 1fr;width:300px;height:80px;background:#fef3c7}" +
            "</style><div class='positioned'>Editable positioned<img src='" + image + "' style='opacity:.4;width:24px;height:18px'></div>" +
            "<div class='grid'>Grid A Grid B</div>";
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using var stream = new MemoryStream();
        result.Value.Save(stream);
        result.Value.Dispose();

        using PowerPointPresentation reopened = PowerPointPresentation.Load(
            new MemoryStream(stream.ToArray()),
            new PowerPointLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly });
        PowerPointSlide slide = Assert.Single(reopened.Slides);
        PowerPointTextBox positioned = Assert.Single(slide.TextBoxes, box => box.Text == "Editable positioned");
        PowerPointTextBox grid = Assert.Single(slide.TextBoxes, box => box.Text == "Grid A Grid B");

        Assert.InRange(positioned.LeftPoints, 59.9D, 60.1D);
        Assert.InRange(positioned.TopPoints, 53.9D, 54.1D);
        Assert.InRange(positioned.WidthPoints, 179.9D, 180.1D);
        Assert.Equal("DBEAFE", positioned.FillColor);
        Assert.Equal("FEF3C7", grid.FillColor);
        Assert.False(Overlaps(positioned, grid));
        Assert.Contains(slide.Pictures, picture => picture.FillTransparency == 60);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.BackgroundLayersFlattened);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.RegionProjected);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified);
        Assert.True(result.Succeeded);
    }

    [Fact]
    public void SingleCssShadowApproximationIsDiagnosed() {
        const string html = "<div style='position:absolute;width:180px;height:50px;"
            + "box-shadow:inset 20px 0 0 rgba(255,0,0,.3)'>Shadowed</div>";

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;

        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.EffectUnsupported
            && diagnostic.Detail == "shadowLayers=1; nativeShadowParameters=approximated");
    }

    [Fact]
    public void InFlowRegionMovesBelowExistingSemanticTextInsteadOfCoveringIt() {
        const string html = "<p>Ordinary flow</p><div style='display:grid;width:300px;height:80px;background:#fef3c7'>Grid content</div>";
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);
        PowerPointTextBox flow = Assert.Single(slide.TextBoxes, box => box.Text == "Ordinary flow");
        PowerPointTextBox grid = Assert.Single(slide.TextBoxes, box => box.Text == "Grid content");

        Assert.True(grid.TopPoints >= flow.TopPoints + flow.HeightPoints);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified);
    }

    [Fact]
    public void InFlowRegionMovesBelowSemanticTableInsteadOfCoveringIt() {
        const string html = "<table><tbody><tr><th>Heading</th></tr><tr><td>Table value</td></tr></tbody></table>"
            + "<div style='display:grid;width:300px;height:80px;background:#fef3c7'>Grid after table</div>";
        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic });
        using PowerPointPresentation presentation = result.Value;
        PowerPointSlide slide = Assert.Single(presentation.Slides);
        PowerPointTable table = Assert.Single(slide.Tables);
        PowerPointTextBox grid = Assert.Single(slide.TextBoxes, box => box.Text == "Grid after table");

        Assert.False(Overlaps(table, grid));
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlEditableLayoutDiagnosticCodes.PlacementSimplified);
    }

    [Fact]
    public void LayoutPicturesRespectSharedImageAndShapeBudgets() {
        string image = "data:image/png;base64," + Convert.ToBase64String(PdfPngTestImages.CreateRgbPng(2, 2));
        string html = "<div style='position:absolute;width:120px;height:60px'>Budgeted region"
            + "<img src='" + image + "' style='width:12px;height:12px'><img src='" + image
            + "' style='width:12px;height:12px'></div>";
        HtmlImportLimits limits = HtmlImportLimits.CreateDefault();
        limits.MaxImages = 1;
        limits.MaxShapes = 3;

        HtmlToPowerPointResult result = HtmlConversionDocument.Parse(html)
            .ToPowerPointPresentationResult(new HtmlToPowerPointOptions { Mode = HtmlImportMode.Generic, Limits = limits });
        using PowerPointPresentation presentation = result.Value;

        Assert.Single(Assert.Single(presentation.Slides).Pictures);
        Assert.Contains(result.Report.Diagnostics, diagnostic =>
            diagnostic.Code == HtmlConversionDiagnosticCodes.TargetLimitExceeded);
    }

    private static bool Overlaps(PowerPointShape first, PowerPointShape second) =>
        first.LeftPoints < second.LeftPoints + second.WidthPoints
        && first.LeftPoints + first.WidthPoints > second.LeftPoints
        && first.TopPoints < second.TopPoints + second.HeightPoints
        && first.TopPoints + first.HeightPoints > second.TopPoints;
}
