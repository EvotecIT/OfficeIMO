using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Moving and resizing shapes on slides without glued connectors writes the same slides without the SDK as through it,
/// and doesn't open it; a glued connector still follows its shape.
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
            source = presentation.ToBytes();
        }

        using var loaded = PowerPointDocument.Load(new MemoryStream(source));
        PowerPointConnectionShape connector = loaded.Slides[0].Shapes.OfType<PowerPointConnectionShape>().Single();
        OfficeLength before = connector.Height;
        loaded.Slides[0].Shapes.OfType<PowerPointAutoShape>().Last().Top = OfficeLength.Points(200);

        Assert.NotEqual(before, connector.Height);
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
