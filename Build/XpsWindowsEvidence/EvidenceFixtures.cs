using System.IO;
using System.Xml.Linq;
using OfficeIMO.Xps;

namespace OfficeIMO.XpsWindowsEvidence;

internal static class EvidenceFixtures {
    internal sealed record Fixture(string Name, XpsDocument Document, byte[]? IndependentPackage = null,
        int PageIndex = 0, string Producer = "OfficeIMO.Xps authored Microsoft XPS package");

    internal static IEnumerable<Fixture> Create(string repository, string output) {
        yield return Stroke("triangle-round", "M40,80L160,80", "Triangle", "Round", "Flat");
        yield return Stroke("returning-endpoints", "M80,80L140,80L80,80", "Triangle", "Round", "Flat");
        yield return Stroke("closed-degenerate", "M80,80L80,80Z", "Flat", "Flat", "Flat");
        yield return Stroke("open-degenerate", "M80,80L80,80", "Triangle", "Round", "Flat");
        yield return Stroke("degenerate-join", "M20,70L70,70L70,70L70,20", "Flat", "Flat", "Flat");
        yield return Stroke("closed-degenerate-seam", "M70,70L70,70L70,20L20,20L20,70Z", "Flat", "Flat", "Flat");
        yield return Stroke("zero-triangle-dashes", "M40,40L40,40L112,136", "Flat", "Flat", "Triangle", "0 1");
        yield return Stroke("zero-square-dashes", "M40,40L64,72L64,72L112,136", "Flat", "Flat", "Square", "0 1");
        yield return Stroke("clipped-miter", "M20,70L70,70L70,20", "Flat", "Flat", "Flat", miter: 1);
        yield return Stroke("mixed-caps-dashes", "M40,80L160,80", "Square", "Triangle", "Round", "1 .5");
        yield return Stroke("transparent-overlap", "M40,80L160,80", "Round", "Triangle", "Round", "1 .25", color: "#800000FF");
        foreach (string spread in new[] { "Pad", "Repeat", "Reflect" }) {
            yield return Radial("radial-interior-" + spread.ToLowerInvariant(), 100, spread);
            yield return Radial("radial-boundary-" + spread.ToLowerInvariant(), 160, spread);
            yield return Radial("radial-exterior-" + spread.ToLowerInvariant(), 200, spread);
        }
        yield return Radial("radial-affine-reflect", 100, "Reflect", "1,.3,.2,1,-20,-15");
        string directory = Path.Combine(repository, "OfficeIMO.Xps.Tests", "Fixtures", "ImageDefaults");
        foreach (string file in File.ReadLines(Path.Combine(directory, "expected.csv")).Skip(1)
            .Select(line => line.Split(',')[0]).Distinct(StringComparer.Ordinal)) {
            var document = XpsDocument.Create(XpsFormat.Xps);
            string resource = document.AddResource("Images/source" + Path.GetExtension(file),
                File.ReadAllBytes(Path.Combine(directory, file)), file.EndsWith(".tif", StringComparison.Ordinal) ? "image/tiff" : "image/jpeg");
            document.AddPage(16, 8).AddImage(resource, 0, 0, 16, 8);
            yield return new("image-" + file.Replace('.', '-'), document);
        }
        byte[] nativePackage = WpfFixtureProducer.Create(repository, output);
        var loaded = XpsDocument.Load(nativePackage);
        if (loaded.Documents.Count != 2 || loaded.Pages.Count != 3)
            throw new InvalidDataException("The WPF producer's two-document, three-page sequence was not retained.");
        for (int page = 0; page < loaded.Pages.Count; page++)
            yield return new("independent-wpf-page-" + (page + 1), loaded, nativePackage, page, "Microsoft WPF XpsDocumentWriter");
    }

    private static Fixture Stroke(string name, string geometry, string start, string end, string dash,
        string? dashes = null, double miter = 10, string color = "#FF0000FF") {
        var document = XpsDocument.Create(XpsFormat.Xps);
        var page = document.AddPage(200, 160).AddPath(geometry, null, color, 20);
        var xml = page.GetMarkup();
        var path = xml.Elements().Single();
        path.SetAttributeValue("StrokeStartLineCap", start);
        path.SetAttributeValue("StrokeEndLineCap", end);
        path.SetAttributeValue("StrokeDashCap", dash);
        path.SetAttributeValue("StrokeMiterLimit", miter);
        if (dashes != null) path.SetAttributeValue("StrokeDashArray", dashes);
        page.ReplaceMarkup(xml);
        return new(name, document);
    }

    private static Fixture Radial(string name, int focus, string spread, string transform = "1,0,0,1,0,0") {
        var document = XpsDocument.Create(XpsFormat.Xps);
        var page = document.AddPage(200, 160);
        var xml = page.GetMarkup();
        XNamespace ns = xml.Name.Namespace;
        xml.Add(new XElement(ns + "Path", new XAttribute("Data", "M0,0H200V160H0Z"),
            new XElement(ns + "Path.Fill", new XElement(ns + "RadialGradientBrush",
                new XAttribute("MappingMode", "Absolute"),
                new XAttribute("Center", "100,80"), new XAttribute("GradientOrigin", focus + ",80"),
                new XAttribute("SpreadMethod", spread), new XAttribute("RadiusX", "60"), new XAttribute("RadiusY", "25"),
                new XAttribute("Transform", transform), new XElement(ns + "RadialGradientBrush.GradientStops",
                    new XElement(ns + "GradientStop", new XAttribute("Offset", "0"), new XAttribute("Color", "#FFFF0000")),
                    new XElement(ns + "GradientStop", new XAttribute("Offset", "1"), new XAttribute("Color", "#FF0000FF")))))));
        page.ReplaceMarkup(xml);
        return new(name, document);
    }
}
