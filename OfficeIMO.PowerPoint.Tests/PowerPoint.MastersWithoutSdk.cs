using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Slide masters, layouts, their placeholders and the slide's layout are read and edited without the SDK, writing the
/// same parts as through it.
/// </summary>
public class PowerPointMastersWithoutSdkTests {
    [Fact]
    public void MasterAndLayoutEditsWriteTheSamePartsWithoutTheSdk() {
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
            Assert.False(presentation.IsSdkLoaded, string.Join("; ", presentation.Diagnostics.Select(diagnostic => diagnostic.Message)));
            withoutSdk = presentation.ToBytes();
        }

        foreach (string prefix in new[] { "ppt/slideMasters/slideMaster", "ppt/slideLayouts/slideLayout", "ppt/slides/slide" }) {
            Assert.Equal(PowerPointTablesWithoutSdkTests.Parts(throughSdk, prefix), PowerPointTablesWithoutSdkTests.Parts(withoutSdk, prefix));
        }
    }

    [Fact]
    public void ReadingMastersAndLayoutsMatchesTheSdk() {
        byte[] edited;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            Edit(presentation);
            edited = presentation.ToBytes();
        }

        using var loaded = PowerPointDocument.Load(new MemoryStream(edited));
        string read = Describe(loaded);
        Assert.False(loaded.IsSdkLoaded);
        _ = loaded.OpenXmlDocument;
        Assert.Equal(Describe(loaded), read);
    }

    [Fact]
    public void AMasterPictureAndBackgroundImageAreAddedWithoutTheSdk() {
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            PowerPointSlideMaster master = presentation.SlideMasters[0];
            master.AddImage(new MemoryStream(CreatePng()), OfficeImageFormat.Png, OfficeLength.Points(10), OfficeLength.Points(10), OfficeLength.Points(40), OfficeLength.Points(40));
            master.Layouts[1].SetBackgroundImage(new MemoryStream(CreatePng()), OfficeImageFormat.Png);
            Assert.False(presentation.IsSdkLoaded);
            saved = presentation.ToBytes();
        }

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        var master2 = reopened.OpenXmlDocument.PresentationPart!.SlideMasterParts.First();
        Assert.Single(master2.ImageParts);
        Assert.Single(master2.SlideLayoutParts.ElementAt(1).ImageParts);
        Assert.IsType<PowerPointImage>(reopened.SlideMasters[0].Shapes.Last());
    }

    [Fact]
    public void ReplacingALayoutBackgroundPictureDropsThePreviousPicture() {
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            PowerPointSlideLayout layout = presentation.SlideMasters[0].Layouts[1];
            layout.SetBackgroundImage(new MemoryStream(CreatePng()), OfficeImageFormat.Png);
            layout.SetBackgroundImage(new MemoryStream(PngWithDifferentPixels()), OfficeImageFormat.Png);
            Assert.False(presentation.IsSdkLoaded);
            saved = presentation.ToBytes();
        }

        using var archive = new ZipArchive(new MemoryStream(saved));
        Assert.Single(archive.Entries, entry => entry.FullName.StartsWith("ppt/media/", StringComparison.Ordinal));
    }

    private static byte[] PngWithDifferentPixels() {
        byte[] rgba = new byte[4 * 4 * 4];
        for (int index = 0; index < rgba.Length; index++) rgba[index] = (byte)(255 - index * 3);
        return OfficeIMO.Drawing.OfficePngWriter.EncodeRgba(4, 4, rgba);
    }

    [Fact]
    public void ALayoutPlaceholderWithoutATextBodyIsStillATextBox() {
        byte[] source = RewritePart(CreateDeck(), "ppt/slideLayouts/slideLayout2.xml", root => {
            XElement title = root.Descendants().First(element => element.Name.LocalName == "ph" && (string?)element.Attribute("type") == "title")
                .Ancestors().First(element => element.Name.LocalName == "sp");
            title.Elements().Where(element => element.Name.LocalName == "txBody").Remove();
        });

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        PowerPointTextBox? title = presentation.GetLayoutPlaceholderTextBox(0, 1, PowerPointPlaceholderType.Title);
        Assert.NotNull(title);
        presentation.SetLayoutPlaceholderBounds(0, 1, PowerPointPlaceholderType.Title,
            new PowerPointLayoutBox(OfficeLength.Points(10), OfficeLength.Points(10), OfficeLength.Points(400), OfficeLength.Points(50)));
        Assert.False(presentation.IsSdkLoaded);
    }

    [Fact]
    public void LayoutInfoReportsTheLayoutRelationshipEvenAfterAnotherToTheSamePart() {
        byte[] source = RewritePart(CreateDeck(), "ppt/slideMasters/_rels/slideMaster1.xml.rels", root => {
            XElement layout = root.Elements().First(element => ((string?)element.Attribute("Type"))?.EndsWith("/slideLayout", StringComparison.Ordinal) == true);
            root.AddFirst(new XElement(root.GetDefaultNamespace() + "Relationship", new XAttribute("Id", "rId900"),
                new XAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"),
                new XAttribute("Target", (string)layout.Attribute("Target")!)));
        });

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        string? read = presentation.GetSlideLayouts()[0].RelationshipId;
        Assert.NotEqual("rId900", read);
        _ = presentation.OpenXmlDocument;
        Assert.Equal(presentation.GetSlideLayouts()[0].RelationshipId, read);
    }

    [Fact]
    public void AnEmptyLayoutPartGetsARootWhenEdited() {
        byte[] source = RewritePart(CreateDeck(), "ppt/slideLayouts/slideLayout2.xml", (string _) => string.Empty);

        using var presentation = PowerPointDocument.Load(new MemoryStream(source));
        presentation.EnsureLayoutFooterPlaceholderTextBox(0, 1, "Footer");
        PowerPointSlideLayout copy = presentation.SlideMasters[0].Layouts[1].Duplicate("Copy");
        Assert.Equal("Copy", copy.Name);
        using var reopened = PowerPointDocument.Load(new MemoryStream(presentation.ToBytes()));
        Assert.Equal("Footer", reopened.GetLayoutPlaceholderTextBox(0, 1, PowerPointPlaceholderType.Footer)!.Text);
    }

    private static byte[] RewritePart(byte[] package, string name, Action<XElement> change) =>
        RewritePart(package, name, (string text) => {
            XElement root = XElement.Parse(text);
            change(root);
            return root.ToString(SaveOptions.DisableFormatting);
        });

    private static byte[] RewritePart(byte[] package, string name, Func<string, string> change) {
        var output = new MemoryStream();
        output.Write(package, 0, package.Length);
        using (var archive = new ZipArchive(output, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry entry = archive.GetEntry(name)!;
            string text;
            using (var reader = new StreamReader(entry.Open(), Encoding.UTF8)) text = reader.ReadToEnd();
            entry.Delete();
            using var writer = new StreamWriter(archive.CreateEntry(name).Open(), new UTF8Encoding(false));
            writer.Write(change(text));
        }
        return output.ToArray();
    }

    private static void Edit(PowerPointDocument presentation) {
        PowerPointSlideMaster master = presentation.SlideMasters[0];
        master.Name = "Corporate";
        master.BackgroundColor = OfficeColor.FromRgb(0xF2, 0xF2, 0xF2);
        master.AddShape(OfficePresetShapeType.Rectangle, OfficeLength.Points(0), OfficeLength.Points(0), OfficeLength.Points(960), OfficeLength.Points(12), "Band");
        master.AddTextBox("Confidential", OfficeLength.Points(20), OfficeLength.Points(500), OfficeLength.Points(200), OfficeLength.Points(20));

        PowerPointSlideLayout layout = master.Layouts[1];
        layout.Name = "Content";
        layout.ShowMasterShapes = false;
        layout.BackgroundColor = OfficeColor.Accent1;
        layout.AddPlaceholder(PowerPointPlaceholderType.Picture, new PowerPointLayoutBox(OfficeLength.Points(500), OfficeLength.Points(120), OfficeLength.Points(300), OfficeLength.Points(200)));
        layout.Placeholders[0].Text = "Title here";
        master.Layouts[2].ClearBackground();

        presentation.SetLayoutPlaceholderBounds(0, 1, PowerPointPlaceholderType.Title,
            new PowerPointLayoutBox(OfficeLength.Points(40), OfficeLength.Points(20), OfficeLength.Points(800), OfficeLength.Points(60)));
        presentation.EnsureLayoutFooterPlaceholderTextBox(0, 1, "Footer text");
        presentation.SetLayoutPlaceholderTextStyle(0, 1, PowerPointPlaceholderType.Title, new PowerPointTextStyle(bold: true, color: OfficeColor.Accent2));

        PowerPointSlide[] slides = presentation.Slides.ToArray();
        slides[1].SetLayout(layout);
        slides[2].SetLayout(PowerPointSlideLayoutType.TitleOnly);
        slides[1].AddTextBox("Bound").BindToPlaceholder(PowerPointPlaceholderType.Body);
    }

    private static string Describe(PowerPointDocument presentation) {
        var builder = new StringBuilder();
        foreach (PowerPointSlideMaster master in presentation.SlideMasters) {
            builder.Append(string.Join(";", master.Index, master.Name, master.BackgroundColor, master.Shapes.Count, master.Placeholders.Count)).Append('\n');
            foreach (PowerPointSlideLayout layout in master.Layouts) {
                builder.Append("  ").Append(string.Join(";", layout.Index, layout.Name, layout.LayoutType, layout.ShowMasterShapes, layout.BackgroundColor,
                    layout.Slides.Count, layout.Master.Index, string.Join(",", layout.Shapes.Select(shape => shape.Name)))).Append('\n');
            }
        }
        foreach (PowerPointSlideLayoutInfo info in presentation.GetSlideLayouts()) builder.Append(string.Join(";", info.LayoutIndex, info.Name, info.Type, info.RelationshipId)).Append('\n');
        foreach (PowerPointSlide slide in presentation.Slides) {
            builder.Append(string.Join(";", slide.LayoutIndex, slide.Layout?.Name,
                string.Join(",", slide.GetLayoutPlaceholders().Select(placeholder => placeholder.Name + ":" + placeholder.PlaceholderType + ":" + placeholder.Bounds?.Left.ToEmus())))).Append('\n');
        }
        builder.Append(presentation.GetLayoutIndex(PowerPointSlideLayoutType.TitleOnly)).Append(';').Append(presentation.GetLayoutIndex("Content"));
        return builder.ToString();
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create();
        for (int index = 0; index < 3; index++) presentation.AddSlide().AddTextBox("Slide " + index);
        return presentation.ToBytes();
    }

    private static byte[] CreatePng() {
        byte[] rgba = new byte[4 * 4 * 4];
        for (int index = 0; index < rgba.Length; index++) rgba[index] = (byte)(index * 7);
        return OfficeIMO.Drawing.OfficePngWriter.EncodeRgba(4, 4, rgba);
    }
}
