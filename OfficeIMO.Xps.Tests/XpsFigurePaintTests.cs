using System.Linq;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class XpsFigurePaintTests {
    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void UnfilledFiguresStillStrokeAndKeepOtherFiguresFilled(XpsFormat format) {
        var doc = Create(format, false); var image = Render(doc);
        Assert.True(image.GetPixel(50,50).R > 240 && image.GetPixel(50,50).B > 240);
        Assert.True(image.GetPixel(20,50).B > 240 && image.GetPixel(20,50).R < 30);
        Assert.True(image.GetPixel(140,50).R > 240 && image.GetPixel(140,50).B < 30);
        Assert.True(doc.Pages[0].ToSvg().IsComplete);
    }

    [Theory]
    [InlineData(XpsFormat.Xps)]
    [InlineData(XpsFormat.OpenXps)]
    public void UnstrokedSegmentsRestartDashesAndDoNotAcquireFigureCaps(XpsFormat format) {
        var doc = Create(format, true); var image = Render(doc);
        Assert.True(image.GetPixel(71,80).R > 240); // no round figure cap at internal break
        Assert.True(image.GetPixel(78,80).B > 240 && image.GetPixel(78,80).R < 30);
        Assert.True(image.GetPixel(89,80).R > 240); // phase restarts at x75
        Assert.True(image.GetPixel(98,80).R < 30);
        Assert.True(image.GetPixel(118,80).R > 240);
    }

    private static OfficeRasterImage Render(XpsDocument doc) {
        Assert.True(OfficeRasterImageDecoder.TryDecode(doc.Pages[0].ExportImage(OfficeImageExportFormat.Png).Bytes, out var image)); return image!;
    }

    internal static XpsDocument Create(XpsFormat format, bool suppressed) {
        var doc = XpsDocument.Create(format); var page = doc.AddPage(200,160);
        var xml = page.GetMarkup(); var ns = xml.Name.Namespace;
        XElement geometry;
        if (suppressed) geometry = new XElement(ns+"PathGeometry", new XElement(ns+"PathFigure", new XAttribute("StartPoint","40,80"),
            new XElement(ns+"PolyLineSegment",new XAttribute("Points","75,80"),new XAttribute("IsStroked","false")),
            new XElement(ns+"PolyLineSegment",new XAttribute("Points","115,80")),
            new XElement(ns+"PolyLineSegment",new XAttribute("Points","160,80"),new XAttribute("IsStroked","false"))));
        else geometry = new XElement(ns+"PathGeometry",
            new XElement(ns+"PathFigure",new XAttribute("StartPoint","20,20"),new XAttribute("IsClosed","true"),new XAttribute("IsFilled","false"),
                new XElement(ns+"PolyLineSegment",new XAttribute("Points","80,20 80,80 20,80"))),
            new XElement(ns+"PathFigure",new XAttribute("StartPoint","110,20"),new XAttribute("IsClosed","true"),
                new XElement(ns+"PolyLineSegment",new XAttribute("Points","170,20 170,80 110,80"))));
        var path = new XElement(ns+"Path",new XAttribute("Fill",suppressed?"#00000000":"#FFFF0000"),new XAttribute("Stroke","#FF0000FF"),
            new XAttribute("StrokeThickness","20"),new XAttribute("StrokeStartLineCap","Round"),new XAttribute("StrokeEndLineCap","Round"),
            new XElement(ns+"Path.Data",geometry));
        if(suppressed) path.SetAttributeValue("StrokeDashArray",".5 .5");
        xml.Add(path); page.ReplaceMarkup(xml); return doc;
    }
}
