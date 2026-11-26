using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Fill, outline, transform and effect edits write the same slides without the SDK as through it, and don't open it.
/// </summary>
public class PowerPointAppearanceWithoutSdkTests {
    [Fact]
    public void AppearanceEditsWriteTheSameSlidesWithoutTheSdk() {
        byte[] source = CreateDeck();

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

        Assert.Equal(Slides(throughSdk), Slides(withoutSdk));
    }

    [Fact]
    public void ReadingAppearanceMatchesTheSdk() {
        byte[] edited;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            Edit(presentation);
            edited = presentation.ToBytes();
        }

        using var withoutSdk = PowerPointDocument.Load(new MemoryStream(edited));
        string[] read = withoutSdk.Slides.SelectMany(slide => slide.Shapes).Select(Describe).ToArray();
        Assert.False(withoutSdk.IsSdkLoaded);
        _ = withoutSdk.OpenXmlDocument;
        Assert.Equal(withoutSdk.Slides.SelectMany(slide => slide.Shapes).Select(Describe).ToArray(), read);
    }

    [Fact]
    public void PercentFormColorValuesReadAsTheirIntegerForm() {
        byte[] source;
        using (var presentation = PowerPointDocument.Create()) {
            presentation.AddSlide().AddShape(OfficePresetShapeType.Rectangle).FillColor = OfficeColor.FromRgb(0x12, 0x34, 0x56);
            source = presentation.ToBytes();
        }
        // Strict documents write percentages with a trailing %; hue stays an angle.
        source = RewriteSlide(source, "<a:srgbClr val=\"123456\"/>", "<a:hslClr hue=\"0\" sat=\"100%\" lum=\"50%\"><a:alpha val=\"50%\" /></a:hslClr>");

        using var loaded = PowerPointDocument.Load(new MemoryStream(source));
        PowerPointShape shape = loaded.Slides[0].Shapes.Single();

        Assert.Equal(50, shape.FillTransparency);
        OfficeColor exported = shape.ResolveFillColorForExport()!.Value;
        Assert.Equal((255, 0, 0, 128), (exported.R, exported.G, exported.B, exported.A));
    }

    private static void Edit(PowerPointDocument presentation) {
        foreach (PowerPointSlide slide in presentation.Slides) {
            PowerPointShape[] shapes = slide.Shapes.ToArray();
            PowerPointShape first = shapes[0], second = shapes[1];
            first.FillColor = OfficeColor.FromRgb(0x1F, 0x4E, 0x79);
            first.FillTransparency = 25;
            first.OutlineColor = OfficeColor.Accent2;
            first.OutlineWidth = OfficeLength.Points(2);
            first.OutlineDash = PowerPointLineDashStyle.Dash;
            first.SetLineEnds(OfficeLineMarkerKind.Triangle, OfficeLineMarkerKind.Oval, PowerPointLineEndWidth.Large);
            first.Rotation = 15;
            first.HorizontalFlip = true;
            first.SetShadow(OfficeColor.FromRgb(0, 0, 0));
            first.SetGlow(OfficeColor.Success, OfficeLength.Points(5));
            first.SetSoftEdges(OfficeLength.Points(2));
            first.SetBlur(OfficeLength.Points(1), grow: true);
            first.SetReflection();
            first.ClearGlow();
            second.FillTransparency = 40;
            second.OutlineTransparency = 30;
            second.VerticalFlip = true;
        }
    }

    private static string Describe(PowerPointShape shape) => string.Join(";",
        shape.FillColor, shape.FillTransparency, shape.OutlineColor, shape.OutlineTransparency, shape.OutlineWidth?.ToEmus(), shape.OutlineDash,
        shape.Rotation, shape.HorizontalFlip, shape.VerticalFlip, shape.ResolveFillColorForExport(), shape.ResolveOutlineColorForExport(), shape.UnsupportedFillKindForExport);

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create();
        for (int index = 0; index < 3; index++) {
            PowerPointSlide slide = presentation.AddSlide();
            slide.AddShape(OfficePresetShapeType.Rectangle);
            slide.AddShape(OfficePresetShapeType.Ellipse);
        }
        return presentation.ToBytes();
    }

    private static byte[] RewriteSlide(byte[] package, string from, string to) {
        var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry("ppt/slides/slide1.xml")!;
            string text;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) text = reader.ReadToEnd();
            Assert.Contains(from, text);
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry("ppt/slides/slide1.xml").Open(), new UTF8Encoding(false));
            writer.Write(text.Replace(from, to));
        }
        return output.ToArray();
    }

    private static string[] Slides(byte[] package) {
        using var archive = new ZipArchive(new MemoryStream(package));
        return archive.Entries.Where(entry => entry.FullName.StartsWith("ppt/slides/slide", StringComparison.Ordinal))
            .OrderBy(entry => entry.FullName, StringComparer.Ordinal)
            .Select(entry => {
                using var reader = new StreamReader(entry.Open(), Encoding.UTF8);
                return Normalize(reader.ReadToEnd());
            }).ToArray();
    }

    // Attribute order and namespace declaration placement aren't part of the content compared.
    private static string Normalize(string markup) {
        XElement element = XElement.Parse(markup);
        foreach (XElement item in element.DescendantsAndSelf()) {
            XAttribute[] attributes = item.Attributes().Where(attribute => !attribute.IsNamespaceDeclaration)
                .OrderBy(attribute => attribute.Name.ToString(), StringComparer.Ordinal).ToArray();
            item.RemoveAttributes();
            item.Add(attributes);
        }
        return element.ToString(SaveOptions.DisableFormatting);
    }
}
