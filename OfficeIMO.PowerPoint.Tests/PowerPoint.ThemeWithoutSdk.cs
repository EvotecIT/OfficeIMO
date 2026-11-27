using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.PowerPoint;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.Tests;

/// <summary>
/// A slide's theme, with the colour scheme of its own theme override, is read from the package without opening the SDK,
/// and reads the same once the SDK is open.
/// </summary>
public class PowerPointThemeWithoutSdkTests {
    [Fact]
    public void ASlideThemeOverrideIsReadFromThePackageAsTheSdkReadsIt() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeckWithRedOverrideOnSecondSlide()));

        OfficeTheme[] withoutSdk = presentation.Slides.Select(slide => slide.AuthoringTheme).ToArray();
        Assert.False(presentation.IsSdkLoaded);
        _ = presentation.OpenXmlDocument;
        OfficeTheme[] withSdk = presentation.Slides.Select(slide => slide.AuthoringTheme).ToArray();

        Assert.Equal(OfficeColor.Parse("FF0000"), withoutSdk[1].Palette.Accent1);
        Assert.NotEqual(withoutSdk[0].Palette.Accent1, withoutSdk[1].Palette.Accent1);
        Assert.Equal(withSdk.Select(theme => theme.Palette.Accent1), withoutSdk.Select(theme => theme.Palette.Accent1));
        Assert.Equal(withSdk.Select(theme => theme.Fonts.Heading.Family), withoutSdk.Select(theme => theme.Fonts.Heading.Family));
        Assert.Equal(withSdk.Select(theme => theme.Name), withoutSdk.Select(theme => theme.Name));
    }

    [Fact]
    public void AThemeColorWithoutASlotIsWrittenWithoutOpeningTheSdk() {
        using var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeckWithRedOverrideOnSecondSlide()));
        PowerPointTextRun run = presentation.Slides[1].TextBoxes.Single().Paragraphs[0].Runs[0];

        run.Color = OfficeColor.Success;

        Assert.False(presentation.IsSdkLoaded);
        Assert.NotNull(run.Color);
    }

    private static byte[] CreateDeckWithRedOverrideOnSecondSlide() {
        byte[] source;
        using (var presentation = PowerPointDocument.Create()) {
            presentation.AddSlide().AddTextBox("First");
            presentation.AddSlide().AddTextBox("Second");
            source = presentation.ToBytes();
        }
        using var stream = new MemoryStream();
        stream.Write(source, 0, source.Length);
        using (PresentationDocument document = PresentationDocument.Open(stream, true)) {
            PresentationPart part = document.PresentationPart!;
            string id = part.Presentation!.SlideIdList!.Elements<P.SlideId>().ElementAt(1).RelationshipId!.Value!;
            var slide = (SlidePart)part.GetPartById(id);
            var colors = (A.ColorScheme)slide.SlideLayoutPart!.SlideMasterPart!.ThemePart!.Theme!.ThemeElements!.ColorScheme!.CloneNode(true);
            colors.Accent1Color = new A.Accent1Color(new A.RgbColorModelHex { Val = "FF0000" });
            slide.AddNewPart<ThemeOverridePart>().ThemeOverride = new A.ThemeOverride(colors);
        }
        return stream.ToArray();
    }
}
