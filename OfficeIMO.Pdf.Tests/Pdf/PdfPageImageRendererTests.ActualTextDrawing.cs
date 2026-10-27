using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfPageImageRendererTests {
    [Fact]
    public void PathActualTextCanvasSurvivesDrawingAndAccessibleSvgWithoutChangingPaint() {
        OfficeShape ink = OfficeShape.Rectangle(30D, 15D);
        ink.FillColor = OfficeColor.Red;
        var options = new PdfCreateOptions { PageWidth = 100D, PageHeight = 80D,
            MarginLeft = 0D, MarginRight = 0D, MarginTop = 0D, MarginBottom = 0D };
        byte[] labelled = PdfDocument.Create(options)
            .Canvas(canvas => canvas
                .Shape(ink, 5D, 5D)
                .ActualText("Outlined AB", 20D, 30D, logical => logical.Shape(ink, 20D, 20D))
                .Shape(ink, 35D, 35D))
            .ToBytes();
        byte[] control = PdfDocument.Create(options)
            .Canvas(canvas => canvas.Shape(ink, 5D, 5D).Shape(ink, 20D, 20D).Shape(ink, 35D, 35D))
            .ToBytes();

        PdfReadPage page = PdfReadDocument.Open(labelled).Pages[0];
        Assert.Equal("Outlined AB", page.ExtractText().Trim());
        OfficeDrawing drawing = page.ToDrawing();
        OfficeDrawingGroup group = Assert.Single(ActualTextGroups(drawing));
        Assert.Equal("Outlined AB", group.ActualText);
        Assert.Single(group.InnerDrawing.Shapes);
        XDocument svg = XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing));
        XElement accessible = Assert.Single(svg.Descendants(), element =>
            (string?)element.Attribute("aria-label") == "Outlined AB");
        Assert.Equal("img", (string?)accessible.Attribute("role"));
        Assert.Contains(accessible.Descendants(), element => element.Name.LocalName is "path" or "rect");
        Assert.Equal(RenderActualTextProof(PdfReadDocument.Open(control).Pages[0].ToDrawing()),
            RenderActualTextProof(drawing));
    }

    private static byte[] RenderActualTextProof(OfficeDrawing drawing) => OfficePngWriter.Encode(
        OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White));

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PathActualTextScopesRetainRepeatedFormOwnership(bool isolated) {
        byte[] pdf = BuildRepeatedActualTextFormPdf(isolated);
        OfficeDrawing drawing = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxActualTextCharacters = 15 }
        }).Pages[0].ToDrawing();

        OfficeDrawingGroup outer = Assert.Single(drawing.Elements.OfType<OfficeDrawingGroup>(),
            group => group.ActualText == "Outer");
        Assert.Equal(new[] { "Inner", "Inner" }, ActualTextGroups(outer.InnerDrawing).Select(group => group.ActualText));
        Assert.Equal(3, XDocument.Parse(OfficeDrawingSvgExporter.ToSvg(drawing)).Descendants()
            .Count(element => element.Attribute("aria-label") != null));
        // Both invocations retain their paint and their own replacement, rather than a cached first invocation.
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        Assert.Equal(OfficeColor.Red, raster.GetPixel(25, 135));
        Assert.Equal(OfficeColor.Red, raster.GetPixel(85, 135));

        PdfReadLimitException limit = Assert.Throws<PdfReadLimitException>(() => PdfReadDocument.Open(pdf,
            new PdfLoadOptions { Limits = new PdfReadLimits { MaxActualTextCharacters = 14 } }).Pages[0].ToDrawing());
        Assert.Equal(PdfReadLimitKind.ActualTextCharacters, limit.Kind);
        Assert.Equal(15, limit.Actual);
    }

    private static byte[] BuildRepeatedActualTextFormPdf(bool isolated) => BuildSingleStreamPdf(
        "/Span << /ActualText (Outer) >> BDC q 1 0 0 1 20 60 cm /Fm Do Q q 1 0 0 1 80 60 cm /Fm Do Q EMC",
        "<< /XObject << /Fm 5 0 R >> >>",
        BuildStreamObject(5, "<< /Type /XObject /Subtype /Form /BBox [0 0 20 20] /Resources << /Properties << /Label << /ActualText (Inner) >> >> >>" +
            (isolated ? " /Group << /S /Transparency /I true /K false >>" : string.Empty),
            "/Span /Label BDC 1 0 0 rg 0 0 20 20 re f EMC"));

    [Theory]
    [InlineData("hidden")]
    [InlineData("artifact")]
    [InlineData("unclosed")]
    [InlineData("empty")]
    [InlineData("visible-text")]
    public void PathActualTextDoesNotRelabelHiddenUnclosedEmptyOrVisibleTextScopes(string kind) {
        const string paint = "1 0 0 rg 20 60 20 20 re f";
        string content = kind switch {
            "hidden" => "/OC /Hidden BDC /Span << /ActualText (Hidden) >> BDC " + paint + " EMC EMC",
            "artifact" => "/Artifact BMC /Span << /ActualText (Decorative) >> BDC " + paint + " EMC EMC",
            "unclosed" => "/Span << /ActualText (Unclosed) >> BDC " + paint,
            "empty" => "/Span << /ActualText (Empty) >> BDC q Q EMC",
            _ => "/Span << /ActualText (Logical) >> BDC BT /F1 12 Tf 1 0 0 1 20 100 Tm (A) Tj ET " + paint + " EMC"
        };
        byte[] pdf = BuildSingleStreamPdf(content,
            "<< /Font << /F1 5 0 R >> /Properties << /Hidden 6 0 R >> >>",
            "5 0 obj\n<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>\nendobj",
            "6 0 obj\n<< /Type /OCG /Name (Hidden) >>\nendobj");
        if (kind == "hidden") {
            pdf = Encoding.ASCII.GetBytes(Encoding.ASCII.GetString(pdf).Replace(
                "<< /Type /Catalog /Pages 2 0 R >>",
                "<< /Type /Catalog /Pages 2 0 R /OCProperties << /OCGs [6 0 R] /D << /OFF [6 0 R] >> >> >>"));
        }
        OfficeDrawing drawing = PdfReadDocument.Open(pdf, new PdfLoadOptions {
            Limits = new PdfReadLimits { MaxActualTextCharacters = 1 }
        }).Pages[0].ToDrawing();

        Assert.Empty(ActualTextGroups(drawing));
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        Assert.Equal(kind is "hidden" or "empty" ? OfficeColor.White : OfficeColor.Red, raster.GetPixel(25, 135));
        if (kind == "visible-text") Assert.Contains(drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "A");
    }

    private static IEnumerable<OfficeDrawingGroup> ActualTextGroups(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingGroup group) {
                if (group.ActualText != null) yield return group;
                foreach (OfficeDrawingGroup child in ActualTextGroups(group.InnerDrawing)) yield return child;
            } else if (element is OfficeDrawingEffectGroup effect) {
                foreach (OfficeDrawingGroup child in ActualTextGroups(effect.InnerDrawing)) yield return child;
            }
        }
    }
}
