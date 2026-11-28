using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Moving and resizing shapes on slides without glued connectors writes the same slides without the SDK as through it,
/// and doesn't open it; glued connectors follow their shapes on the owned package.
/// </summary>
public class PowerPointGeometryWithoutSdkTests {
    [Fact]
    public void MovingAndResizingShapesWritesTheSameSlidesWithoutTheSdk() {
        byte[] source;
        using (var presentation = PowerPointDocument.Create()) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddShape(OfficePresetShapeType.Rectangle);
            slide.AddTextBox("Text");
            source = presentation.ToBytes();
        }

        byte[] throughSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            _ = presentation.OpenXmlDocument;
            Edit(presentation);
            throughSdk = presentation.ToBytes();
        }
        byte[] withoutSdk;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            Edit(presentation);
            Assert.False(presentation.IsSdkLoaded);
            withoutSdk = presentation.ToBytes();
        }

        Assert.Equal(Slide(throughSdk), Slide(withoutSdk));
    }

    [Fact]
    public void AGluedConnectorFollowsAShapeMovedAfterLoading() {
        byte[] source;
        using (var presentation = PowerPointDocument.Create()) {
            PowerPointSlide slide = presentation.AddSlide();
            PowerPointShape from = slide.AddShape(OfficePresetShapeType.Rectangle, OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(80), OfficeLength.Points(40));
            PowerPointShape to = slide.AddShape(OfficePresetShapeType.Rectangle, OfficeLength.Points(300), OfficeLength.Points(20), OfficeLength.Points(80), OfficeLength.Points(40));
            slide.AddConnector(from, to, PowerPointConnectorKind.Straight);
            Assert.False(presentation.IsSdkLoaded);
            source = presentation.ToBytes();
        }

        using var loaded = PowerPointDocument.Load(new MemoryStream(source));
        PowerPointConnectionShape connector = loaded.Slides[0].Shapes.OfType<PowerPointConnectionShape>().Single();
        OfficeLength before = connector.Height;
        loaded.Slides[0].Shapes.OfType<PowerPointAutoShape>().Last().Top = OfficeLength.Points(200);

        Assert.NotEqual(before, connector.Height);
        Assert.False(loaded.IsSdkLoaded);
    }

    [Fact]
    public void UngroupedConnectorsResumeFollowingTheirEndpoints() {
        using var original = PowerPointDocument.Create();
        PowerPointSlide slide = original.AddSlide();
        PowerPointAutoShape from = slide.AddRectangle(OfficeLength.Points(30), OfficeLength.Points(35),
            OfficeLength.Points(80), OfficeLength.Points(40), "From");
        PowerPointAutoShape to = slide.AddRectangle(OfficeLength.Points(280), OfficeLength.Points(35),
            OfficeLength.Points(80), OfficeLength.Points(40), "To");
        PowerPointConnectionShape connector = slide.AddConnector(from, to, PowerPointConnectorKind.Straight);
        slide.GroupShapes(new PowerPointShape[] { from, to, connector });
        slide.AddShape(OfficePresetShapeType.Ellipse, name: "Unrelated");

        using var reopened = PowerPointDocument.Load(new MemoryStream(original.ToBytes()));
        PowerPointSlide loaded = reopened.Slides.Single();
        loaded.GetShapeByName("Unrelated")!.Top = OfficeLength.Points(90);
        loaded.UngroupShape(loaded.Shapes.OfType<PowerPointGroupShape>().Single());
        PowerPointConnectionShape ungrouped = loaded.Shapes.OfType<PowerPointConnectionShape>().Single();
        OfficeLength before = ungrouped.Height;

        loaded.GetShapeByName("To")!.Top = OfficeLength.Points(200);

        Assert.NotEqual(before, ungrouped.Height);
        Assert.False(reopened.IsSdkLoaded);
    }

    [Fact]
    public void ConnectorEndpointsIgnoreMatchingElementsInExtensions() {
        using var source = PowerPointDocument.Create();
        PowerPointSlide slide = source.AddSlide();
        PowerPointAutoShape target = slide.AddShape(OfficePresetShapeType.Rectangle);
        slide.AddConnector(OfficeLength.Points(20), OfficeLength.Points(20), OfficeLength.Points(80), OfficeLength.Points(60),
            PowerPointConnectorKind.Straight);
        using var original = new ZipArchive(new MemoryStream(source.ToBytes()), ZipArchiveMode.Read);
        using var output = new MemoryStream();
        XNamespace p = "http://schemas.openxmlformats.org/presentationml/2006/main";
        XNamespace a = "http://schemas.openxmlformats.org/drawingml/2006/main";
        using (var fixture = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry entry in original.Entries) {
                using Stream input = entry.Open();
                using Stream result = fixture.CreateEntry(entry.FullName).Open();
                if (entry.FullName != "ppt/slides/slide1.xml") {
                    input.CopyTo(result);
                    continue;
                }
                XDocument markup = XDocument.Load(input);
                markup.Descendants(p + "cNvCxnSpPr").Single().Add(
                    new XElement(a + "extLst", new XElement(a + "ext", new XAttribute("uri", "{AC22EBDA-C2BB-4607-B18D-AD8AA0088DB5}"),
                        new XElement(a + "stCxn", new XAttribute("id", target.Id!.Value), new XAttribute("idx", 0)))));
                markup.Save(result);
            }
        }

        output.Position = 0;
        using var loaded = PowerPointDocument.Load(output);
        PowerPointConnectionShape connector = loaded.Slides.Single().Shapes.OfType<PowerPointConnectionShape>().Single();

        Assert.Null(connector.StartShape);
        Assert.Null(connector.StartSite);
        Assert.False(loaded.IsSdkLoaded);
    }

    private static void Edit(PowerPointDocument presentation) {
        foreach (PowerPointShape shape in presentation.Slides[0].Shapes) {
            shape.Left = OfficeLength.Points(40);
            shape.Top = OfficeLength.Points(60);
            shape.Width = OfficeLength.Points(200);
            shape.Height = OfficeLength.Points(80);
        }
    }

    private static string Slide(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package));
        using var reader = new StreamReader(archive.GetEntry("ppt/slides/slide1.xml")!.Open(), Encoding.UTF8);
        XElement element = XElement.Parse(reader.ReadToEnd());
        foreach (XElement item in element.DescendantsAndSelf()) {
            XAttribute[] attributes = item.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToArray();
            item.RemoveAttributes();
            item.Add(attributes);
        }
        return element.ToString(SaveOptions.DisableFormatting);
    }
}
