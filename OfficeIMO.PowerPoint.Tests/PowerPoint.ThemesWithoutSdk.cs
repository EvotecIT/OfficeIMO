using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// Theme colours, fonts and names are read and edited without the SDK, writing the same theme parts as through it.
/// </summary>
public class PowerPointThemesWithoutSdkTests {
    [Fact]
    public void ThemeEditsWriteTheSameThemePartsWithoutTheSdk() {
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

        Assert.Equal(PowerPointTablesWithoutSdkTests.Parts(throughSdk, "ppt/theme/theme"), PowerPointTablesWithoutSdkTests.Parts(withoutSdk, "ppt/theme/theme"));
    }

    [Fact]
    public void ReadingThemesMatchesTheSdk() {
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
        Assert.Contains("Accent1=1F4E79", read);
        Assert.Contains("Major=Segoe UI Semibold", read);
    }

    [Fact]
    public void AThemeFontGetsTheClassificationOfAKnownTypeface() {
        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(CreateDeck()))) {
            presentation.SetThemeLatinFonts("Arial", "Times New Roman");
            Assert.False(presentation.IsSdkLoaded);
            saved = presentation.ToBytes();
        }

        string theme = PartText(saved, "ppt/theme/theme1.xml");
        Assert.Contains("<a:latin typeface=\"Arial\" panose=\"020B0604020202020204\"", theme);
        Assert.Contains("<a:latin typeface=\"Times New Roman\" panose=\"02020603050405020304\"", theme);
    }

    [Fact]
    public void AMasterWithoutAThemeGetsOneWhenAColourIsSet() {
        byte[] source = WithoutThemeOnSecondMaster(CreateDeck());

        byte[] saved;
        using (var presentation = PowerPointDocument.Load(new MemoryStream(source))) {
            Assert.Null(presentation.GetThemeColor(PowerPointThemeColor.Accent1, 1));
            presentation.SetThemeColor(PowerPointThemeColor.Accent1, "C00000", 1);
            saved = presentation.ToBytes();
        }

        using var reopened = PowerPointDocument.Load(new MemoryStream(saved));
        Assert.Equal("C00000", reopened.GetThemeColor(PowerPointThemeColor.Accent1, 1));
    }

    private static void Edit(PowerPointDocument presentation) {
        presentation.SetThemeColor(PowerPointThemeColor.Accent1, "1F4E79");
        presentation.SetThemeColors(new Dictionary<PowerPointThemeColor, string> {
            [PowerPointThemeColor.Hyperlink] = "0563C1",
            [PowerPointThemeColor.FollowedHyperlink] = "954F72",
            [PowerPointThemeColor.Dark2] = "44546A"
        });
        presentation.SetThemeColorForAllMasters(PowerPointThemeColor.Accent6, "70AD47");
        presentation.SetThemeFonts(new PowerPointThemeFontSet("Segoe UI Semibold", "Segoe UI", "Yu Gothic", null, null, "Arial"));
        presentation.SetThemeLatinFontsForAllMasters("Aptos Display", "Aptos");
        presentation.SetThemeFonts(new PowerPointThemeFontSet("Segoe UI Semibold", null, null, null, null, null));
        presentation.SetThemeName("Brand");
        presentation.SetThemeNameForAllMasters("Brand family");
        presentation.SetThemeName("Brand", 0);
    }

    private static string Describe(PowerPointDocument presentation) {
        var text = new StringBuilder();
        text.Append("Name=").AppendLine(presentation.ThemeName);
        int masters = presentation.SlideMasters.Count;
        for (int index = 0; index < masters; index++) {
            foreach (KeyValuePair<PowerPointThemeColor, string> color in presentation.GetThemeColors(index).OrderBy(pair => pair.Key)) {
                text.Append(index).Append(' ').Append(color.Key).Append('=').AppendLine(color.Value);
            }
            PowerPointThemeFontSet fonts = presentation.GetThemeFonts(index);
            text.Append(index).Append(" Major=").Append(fonts.MajorLatin).Append(" Minor=").Append(fonts.MinorLatin)
                .Append(" EA=").Append(fonts.MajorEastAsian).Append('/').Append(fonts.MinorEastAsian)
                .Append(" CS=").Append(fonts.MajorComplexScript).Append('/').AppendLine(fonts.MinorComplexScript);
            text.Append(index).Append(" Latin=").AppendLine(presentation.GetThemeLatinFonts(index).ToString());
        }
        foreach (PowerPointTableStyleInfo style in presentation.TableStyles) {
            text.Append("Table ").Append(style.StyleId).Append(' ').AppendLine(style.Name);
        }
        return text.ToString();
    }

    private static byte[] CreateDeck() {
        using var presentation = PowerPointDocument.Create();
        presentation.AddSlide().AddTextBox("Themed");
        presentation.AddSlideMaster("Second");
        return presentation.ToBytes();
    }

    private static byte[] WithoutThemeOnSecondMaster(byte[] source) {
        using var stream = new MemoryStream();
        stream.Write(source, 0, source.Length);
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            ZipArchiveEntry rels = archive.GetEntry("ppt/slideMasters/_rels/slideMaster2.xml.rels")!;
            string xml;
            using (var reader = new StreamReader(rels.Open())) xml = reader.ReadToEnd();
            XDocument document = XDocument.Parse(xml);
            document.Root!.Elements().Where(element => ((string?)element.Attribute("Type"))?.EndsWith("/theme", StringComparison.Ordinal) == true).Remove();
            rels.Delete();
            using var writer = new StreamWriter(archive.CreateEntry("ppt/slideMasters/_rels/slideMaster2.xml.rels").Open());
            writer.Write(document.ToString(SaveOptions.DisableFormatting));
        }
        return stream.ToArray();
    }

    private static string PartText(byte[] package, string name) {
        using var archive = new ZipArchive(new MemoryStream(package));
        using var reader = new StreamReader(archive.GetEntry(name)!.Open());
        return reader.ReadToEnd();
    }
}
