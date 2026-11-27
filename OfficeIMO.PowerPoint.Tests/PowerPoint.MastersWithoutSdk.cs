using System.IO.Compression;
using System.Text;
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
