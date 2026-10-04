using System;
using System.Linq;

using System.Xml.Linq;
using OfficeIMO.Drawing;

namespace OfficeIMO.Xps.Tests;

internal static class XpsBrushFixtures {
    internal static XpsDocument Create(string scenario, XpsFormat format = XpsFormat.Xps) {
        if (scenario == "mask-group-opacity") {
            var masked = XpsDocument.Create(format); var maskedPage = masked.AddPage(120, 100);
            var markup = maskedPage.GetMarkup(); var n = markup.Name.Namespace;
            var red = new XElement(n + "Path", new XAttribute("Data", "M10,10H100V90H10Z"), new XAttribute("Fill", "#FFFF0000"));
            markup.Add(new XElement(n + "Canvas", new XAttribute("Opacity", "0.5"), new XAttribute("OpacityMask", "#80000000"), red, new XElement(red)));
            maskedPage.ReplaceMarkup(markup); return masked;
        }
        if (scenario == "visual-transformed" || scenario == "image-mask" || scenario == "visual-mask") {
            var nested = Create(scenario == "image-mask" ? "image-Tile" : "visual-Tile", format);
            var nestedPage = nested.Pages[0]; var markup = nestedPage.GetMarkup(); var n = markup.Name.Namespace;
            var visual = markup.Elements().Single(); var fill = visual.Element(n + "Path.Fill")!;
            if (scenario == "visual-transformed") fill.Elements().Single().SetAttributeValue("Transform", "1,0,0,1,5,7");
            else {
                fill.Name = n + "Path.OpacityMask";
                visual.SetAttributeValue("Fill", "#FFFF0000");
                fill.Elements().Single().SetAttributeValue("Opacity", "0.5");
            }
            nestedPage.ReplaceMarkup(markup); return nested;
        }
        var doc = XpsDocument.Create(format); var page = doc.AddPage(120, 100); var xml = page.GetMarkup(); XNamespace ns = xml.Name.Namespace;
        XElement Path(string data, string fill) => new(ns + "Path", new XAttribute("Data", data), new XAttribute("Fill", fill));
        XElement Gradient() => new(ns + "LinearGradientBrush", new XAttribute("StartPoint", "0,0"), new XAttribute("EndPoint", "40,0"),
            new XElement(ns + "LinearGradientBrush.GradientStops", new XElement(ns + "GradientStop", new XAttribute("Offset", "0"), new XAttribute("Color", "#FFFF0000")), new XElement(ns + "GradientStop", new XAttribute("Offset", "1"), new XAttribute("Color", "#FF0000FF"))));
        var path = Path("M0,0H120V100H0Z", "#FFFF0000");
        if (scenario == "gradient") {
            path.Attribute("Fill")!.Remove(); var gradient = Gradient(); gradient.SetAttributeValue("Transform", "1,0.5,0,1,20,0"); path.Add(new XElement(ns + "Path.Fill", gradient));
        } else if (scenario.StartsWith("image-", StringComparison.Ordinal) || scenario.StartsWith("visual-", StringComparison.Ordinal)) {
            bool image = scenario.StartsWith("image-", StringComparison.Ordinal);
            var brush = new XElement(ns + (image ? "ImageBrush" : "VisualBrush"), new XAttribute("Viewbox", "0,0,4,4"), new XAttribute("Viewport", "10,10,20,20"), new XAttribute("ViewboxUnits", "Absolute"), new XAttribute("ViewportUnits", "Absolute"), new XAttribute("TileMode", scenario.Substring(scenario.IndexOf('-') + 1)));
            if (image) {
                var pixels = new OfficeRasterImage(4, 4, OfficeColor.Red);
                for (int y = 0; y < 4; y++) for (int x = 2; x < 4; x++) pixels.SetPixel(x, y, OfficeColor.Blue);
                for (int y = 2; y < 4; y++) for (int x = 0; x < 4; x++) pixels.SetPixel(x, y, OfficeColor.Green);
                string uri = doc.AddResource("Resources/tile.png", OfficePngWriter.Encode(pixels), "image/png"); brush.SetAttributeValue("ImageSource", uri);
            } else brush.Add(new XElement(ns + "VisualBrush.Visual", new XElement(ns + "Canvas", Path("M0,0H2V4H0Z", "#FFFF0000"), Path("M2,0H4V4H2Z", "#FF0000FF"), Path("M0,2H4V4H0Z", "#FF008000"))));
            path.Attribute("Fill")!.Remove(); path.Add(new XElement(ns + "Path.Fill", brush));
        } else if (scenario == "mask" || scenario == "mask-transformed") {
            var gradient = Gradient(); gradient.SetAttributeValue("EndPoint", "100,0");
            var stops = gradient.Element(ns + "LinearGradientBrush.GradientStops")!;
            stops.Elements().First().SetAttributeValue("Color", "#00000000"); stops.Elements().Last().SetAttributeValue("Color", "#FF000000");
            if (scenario == "mask-transformed") {
                path.SetAttributeValue("Data", "M-100,0H20V100H-100Z"); gradient.SetAttributeValue("StartPoint", "-100,0"); gradient.SetAttributeValue("EndPoint", "0,0");
                path = new XElement(ns + "Canvas", new XAttribute("RenderTransform", "1,0,0,1,100,0"), path);
            }
            path.Add(new XElement(ns + (path.Name.LocalName + ".OpacityMask"), gradient));
        } else throw new ArgumentException("Unknown brush scenario", nameof(scenario));
        xml.Add(path); page.ReplaceMarkup(xml); return doc;
    }
}
