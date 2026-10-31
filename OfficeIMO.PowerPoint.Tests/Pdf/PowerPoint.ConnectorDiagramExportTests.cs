using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Html;
using OfficeIMO.PowerPoint.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

/// <summary>
/// An architecture diagram of themed boxes, rotated labels, arrow connectors and a scaled group, on a
/// title-slide layout whose title and subtitle placeholders the slide leaves unused.
/// </summary>
public sealed class PowerPointConnectorDiagramExportTests {
    const string Namespaces =
        " xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"" +
        " xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"";

    // Sample points in slide points, with whether the slide paints there. Arrow samples sit
    // beside the 6pt stroke, so only an arrow head can paint them.
    static readonly (string Name, double X, double Y, bool Painted)[] samples = {
        ("themed box fill", 190, 160, true),
        ("rotated label fill", 22, 125, true),
        ("vertical connector stroke", 380, 200, true),
        ("vertical connector head arrow", 388, 128, true),
        ("vertical connector plain tail", 388, 290, false),
        ("double connector stroke", 180, 330, true),
        ("double connector start arrow", 84, 338, true),
        ("double connector end arrow", 276, 338, true),
        ("group box fill", 540, 150, true),
        ("group connector stroke", 540, 265, true),
        ("group connector head arrow", 547, 246, true)
    };

    [Fact]
    public void PdfDrawsThemedShapesConnectorsAndGroupsLikeTheImageRenderer() {
        using var presentation = CreateDiagram();
        PowerPointSlide slide = presentation.Slides[0];

        PdfDocumentConversionResult result = presentation.ConvertToPdfWithEvidence();
        byte[] pdf = result.ToBytes();

        Assert.DoesNotContain(result.Warnings, _ => _.Code is "unsupported-shape" or "invalid-shape-bounds" or "unsupported-connector");
        PdfReadDocument read = PdfReadDocument.Open(pdf);
        string text = read.ExtractText();
        Assert.Contains("User Experience", text, StringComparison.Ordinal);
        Assert.DoesNotContain("Click to edit", text, StringComparison.Ordinal);

        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(read.Pages[0].ToDrawing(), 1D, OfficeColor.White);
        PowerPointSlideVisualInfo snapshot = slide.CreateVisualSnapshot(new PowerPointImageExportOptions());
        OfficeRasterImage image = OfficeDrawingRasterRenderer.Render(snapshot.Drawing, 1D, OfficeColor.White);
        foreach ((string name, double x, double y, bool painted) in samples) {
            OfficeColor pdfPixel = rendered.GetPixel((int)x, (int)y);
            OfficeColor imagePixel = image.GetPixel((int)x, (int)y);
            Assert.True(IsWhite(pdfPixel) != painted, $"{name} in PDF: observed {pdfPixel}.");
            Assert.True(IsWhite(imagePixel) != painted, $"{name} in image: observed {imagePixel}.");
            Assert.True(Math.Abs(pdfPixel.R - imagePixel.R) <= 8 && Math.Abs(pdfPixel.G - imagePixel.G) <= 8 &&
                Math.Abs(pdfPixel.B - imagePixel.B) <= 8,
                $"{name}: PDF {pdfPixel} differs from image {imagePixel}.");
        }
    }

    [Fact]
    public void VisualHtmlDrawsTheDiagramWithoutPlaceholderPrompts() {
        using var presentation = CreateDiagram();
        PowerPointHtmlSaveOptions options = PowerPointHtmlSaveOptions.CreateVisualReviewProfile();
        options.IncludeMasterInventory = false;
        options.IncludeExtractionProof = false;

        string html = presentation.ToHtml(options);

        Assert.Contains("User Experience", html, StringComparison.Ordinal);
        Assert.DoesNotContain("Click to edit", html, StringComparison.Ordinal);
        // Read the visible snapshot, rather than the hidden positioned-shape metadata.
        int snapshotStart = html.IndexOf("<div class=\"officeimo-shared-slide-snapshot\"", StringComparison.Ordinal);
        Assert.True(snapshotStart >= 0, "Visual HTML must contain the slide snapshot.");
        int svgStart = html.IndexOf("<svg ", snapshotStart, StringComparison.Ordinal);
        Assert.True(svgStart >= 0, "The slide snapshot must contain SVG.");
        int svgEnd = html.IndexOf("</svg>", svgStart, StringComparison.Ordinal);
        Assert.True(svgEnd >= 0, "The slide snapshot SVG must be complete.");
        string svg = html.Substring(svgStart, svgEnd + "</svg>".Length - svgStart);
        Assert.Contains("User Experience", svg, StringComparison.Ordinal);
        Assert.True(OfficeSvgDrawingReader.TryRead(Encoding.UTF8.GetBytes(svg), out OfficeDrawing? drawing),
            "The slide snapshot SVG must be readable as a drawing.");

        // SVG point units may become CSS pixels on import. Sample in the authored slide coordinates.
        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing!,
            presentation.SlideSize.Width.ToPoints() / drawing!.Width, OfficeColor.White);
        foreach ((string name, double x, double y, bool painted) in samples) {
            OfficeColor pixel = rendered.GetPixel((int)x, (int)y);
            Assert.True(IsWhite(pixel) != painted, $"{name} in visual HTML: observed {pixel}.");
        }
    }

    static bool IsWhite(OfficeColor color) => color.R > 245 && color.G > 245 && color.B > 245;

    static PowerPointDocument CreateDiagram() {
        var presentation = PowerPointDocument.Create();
        presentation.SlideSize.SetSize(OfficeIMO.OfficeLength.Points(720),OfficeIMO.OfficeLength.Points(540));
        PowerPointSlide slide = presentation.AddSlide(PowerPointSlideLayoutType.Title);
        // PowerPoint authors these editing prompts into the layout placeholders.
        foreach (Shape placeholder in slide.SlidePart.SlideLayoutPart!.SlideLayout.CommonSlideData!.ShapeTree!.Elements<Shape>()) {
            PlaceholderValues? type = placeholder.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value;
            string prompt = type == PlaceholderValues.SubTitle ? "Click to edit Master subtitle style" : "Click to edit Master title style";
            placeholder.TextBody = new TextBody(
                "<p:txBody" + Namespaces + "><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang=\"en-US\"/><a:t>" + prompt + "</a:t></a:r></a:p></p:txBody>");
        }
        ShapeTree tree = slide.SlidePart.Slide.CommonSlideData!.ShapeTree!;
        foreach (OpenXmlElement shape in tree.ChildElements.Where(_ => _ is Shape or GroupShape or ConnectionShape).ToArray()) {
            shape.Remove();
        }

        tree.Append(Box(10, "Services", 40, 120, 300, 80, "accent6"));
        tree.Append(Label(11, "User Experience", 30, 160, 110, 20));
        tree.Append(Connector(12, "straightConnector1", 380, 100, 0, 200, head: true, tail: false));
        tree.Append(Connector(13, "line", 60, 330, 240, 0, head: true, tail: true));
        // Children are authored at twice the group size, away from the origin, so they
        // only land on the slide through the group's child offset and extent.
        tree.Append(new GroupShape(
            "<p:grpSp" + Namespaces + "><p:nvGrpSpPr><p:cNvPr id=\"20\" name=\"Group 20\"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>" +
            "<p:grpSpPr><a:xfrm><a:off x=\"" + Emu(420) + "\" y=\"" + Emu(120) + "\"/><a:ext cx=\"" + Emu(240) + "\" cy=\"" + Emu(160) + "\"/>" +
            "<a:chOff x=\"" + Emu(1000) + "\" y=\"" + Emu(500) + "\"/><a:chExt cx=\"" + Emu(480) + "\" cy=\"" + Emu(320) + "\"/></a:xfrm></p:grpSpPr>" +
            Box(21, "Data", 1000, 500, 480, 120, "accent3").OuterXml +
            Connector(22, "straightConnector1", 1240, 700, 0, 120, head: true, tail: false).OuterXml +
            "</p:grpSp>"));
        return presentation;
    }

    static Shape Box(uint id, string name, double x, double y, double width, double height, string accent) =>
        new(
            "<p:sp" + Namespaces + "><p:nvSpPr><p:cNvPr id=\"" + id + "\" name=\"" + name + "\"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr>" + Transform(x, y, width, height, rotated: false) + "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom>" +
            "<a:solidFill><a:schemeClr val=\"" + accent + "\"><a:lumMod val=\"20000\"/><a:lumOff val=\"80000\"/></a:schemeClr></a:solidFill>" +
            "<a:ln w=\"9525\"><a:solidFill><a:schemeClr val=\"" + accent + "\"><a:lumMod val=\"40000\"/><a:lumOff val=\"60000\"/></a:schemeClr></a:solidFill></a:ln></p:spPr>" +
            "<p:style><a:lnRef idx=\"2\"><a:schemeClr val=\"accent1\"><a:shade val=\"50000\"/></a:schemeClr></a:lnRef><a:fillRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:fillRef>" +
            "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"lt1\"/></a:fontRef></p:style>" +
            "<p:txBody><a:bodyPr rtlCol=\"0\" anchor=\"ctr\"/><a:lstStyle/><a:p><a:pPr algn=\"ctr\"/><a:endParaRPr lang=\"en-AU\"/></a:p></p:txBody></p:sp>");

    // A label centred on (centerX, centerY), turned to read bottom to top.
    static Shape Label(uint id, string text, double centerX, double centerY, double width, double height) =>
        new(
            "<p:sp" + Namespaces + "><p:nvSpPr><p:cNvPr id=\"" + id + "\" name=\"TextBox " + id + "\"/><p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr>" +
            "<p:spPr>" + Transform(centerX - width / 2, centerY - height / 2, width, height, rotated: true) +
            "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val=\"accent1\"/></a:solidFill></p:spPr>" +
            "<p:txBody><a:bodyPr wrap=\"square\" rtlCol=\"0\"/><a:lstStyle/><a:p><a:pPr algn=\"ctr\"/>" +
            "<a:r><a:rPr lang=\"en-AU\" sz=\"1100\"><a:solidFill><a:schemeClr val=\"bg1\"/></a:solidFill></a:rPr><a:t>" + text + "</a:t></a:r></a:p></p:txBody></p:sp>");

    static ConnectionShape Connector(uint id, string preset, double x, double y, double width, double height, bool head, bool tail) =>
        new(
            "<p:cxnSp" + Namespaces + "><p:nvCxnSpPr><p:cNvPr id=\"" + id + "\" name=\"Connector " + id + "\"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>" +
            "<p:spPr>" + Transform(x, y, width, height, rotated: false) + "<a:prstGeom prst=\"" + preset + "\"><a:avLst/></a:prstGeom>" +
            "<a:ln w=\"76200\"><a:solidFill><a:schemeClr val=\"accent5\"><a:lumMod val=\"75000\"/></a:schemeClr></a:solidFill>" +
            "<a:headEnd type=\"" + (head ? "triangle" : "none") + "\" w=\"med\" len=\"med\"/><a:tailEnd type=\"" + (tail ? "triangle" : "none") + "\" w=\"med\" len=\"med\"/></a:ln></p:spPr>" +
            "<p:style><a:lnRef idx=\"1\"><a:schemeClr val=\"accent1\"/></a:lnRef><a:fillRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:fillRef>" +
            "<a:effectRef idx=\"0\"><a:schemeClr val=\"accent1\"/></a:effectRef><a:fontRef idx=\"minor\"><a:schemeClr val=\"tx1\"/></a:fontRef></p:style></p:cxnSp>");

    static string Transform(double x, double y, double width, double height, bool rotated) =>
        "<a:xfrm" + (rotated ? " rot=\"16200000\"" : string.Empty) + "><a:off x=\"" + Emu(x) + "\" y=\"" + Emu(y) + "\"/>" +
        "<a:ext cx=\"" + Emu(width) + "\" cy=\"" + Emu(height) + "\"/></a:xfrm>";

    static string Emu(double points) => PowerPointUnits.FromPoints(points).ToString(System.Globalization.CultureInfo.InvariantCulture);
}
