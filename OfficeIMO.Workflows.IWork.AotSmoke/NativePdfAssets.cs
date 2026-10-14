using System.Security.Cryptography;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows.IWork;

internal static class NativePdfAssets {
    internal sealed record Source(string Name, string Hash, string Route, int Pages);

    internal static readonly Source Body = new("pages-created-edited-v15.4.pages",
        "DC07FA4028AC9F261FDB7CB23F8B409CD80E092CDF8A8ED938ADC420B37AD6B4", "pages-pdf", 2);
    internal static readonly Source List = new("pages-mixed-size-edited-v15.4.pages",
        "6511DF28A9EB82C1A3F437165395F4541D07AFECDCF96E1B8256F3793E15F9D8", "pages-pdf", 1);
    internal static readonly Source Slides = new("keynote-preserved-paragraphs-edited-v15.4.key",
        "9FF8299C1ACFA80055368F0B541544273087B747E51D1F6AE43E1AB4D8771212", "keynote-pdf", 3);
    internal static readonly Source NativeNumbers = new("numbers-created-edited-v14.5.numbers",
        "C907C567584786CE9C5183A376F61D813F3BF56757DE531CBDBFB1FEFB4FCFD6", "numbers-pdf", 0);
    internal static readonly Source UnpaddedControl = new("numbers-unpadded-scalar.numbers",
        "95878A6826028FD6F49BADFF7885436FE78734A768212694539B9655F679D9F3", "numbers-pdf", 1);

    internal static string PathFor(Source source) {
        string path = Path.Combine(AppContext.BaseDirectory, "Corpus", source.Name);
        Require(Hash(path) == source.Hash, source.Name + " fixture changed.");
        return path;
    }

    internal static IWorkWorkflowSettings PortableSettings() {
        var fonts = new OfficeFontFaceCollection();
        foreach (var (name, style, hash) in new[] {
            ("Regular", OfficeFontStyle.Regular, "F6418F708BAEDE9789DAEF5D458C0F53D2A888AF9820E8062934E504FEDC6595"),
            ("Bold", OfficeFontStyle.Bold, "BB5D20F79B82599EC72983597437373A80F2D2085FA91FC144FD74E876A594DB"),
            ("Italic", OfficeFontStyle.Italic, "0B019225E58D702BFEDCBD35C21696769F8EE115CB6343F84C2F240312450D1C"),
            ("BoldItalic", OfficeFontStyle.Bold | OfficeFontStyle.Italic, "B32928186C119599E03CA6A1FFC680FDCB7FAC95772F4B95D989CF6CD3861517")
        }) fonts.Add("Carlito", Font("Carlito-" + name + ".ttf", hash), style);
        fonts.Add("Noto Emoji", Font("NotoEmoji-VariableFont_wght.ttf", "DBD8394B0EE513B356B5A101FB0F64628CE9508294523E051A0EFB1679BB34E3"))
            .AddFallbackFamily("Noto Emoji");
        fonts.Add("Noto JP Common", Font("NotoSansJP-OfficeIMO-Common.ttf", "4777E0AC6A42D1BAB96F53BD9F87D5D003BF9852CEA1A8358463696D9ED42337"))
            .AddFallbackFamily("Noto JP Common");
        fonts.AddFallbackFamily("Carlito");
        var profile = new OfficeRenderingProfile("iwork-pdf-native-aot", fonts);
        var settings = new IWorkWorkflowSettings();
        settings.PagesPdfOptions.UseRenderingProfile(profile);
        settings.NumbersPdfOptions.UseRenderingProfile(profile);
        settings.KeynotePdfOptions.UseRenderingProfile(profile);
        settings.PagesPdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
        settings.NumbersPdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
        settings.KeynotePdfOptions.ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic();
        return settings;
    }

    private static byte[] Font(string name, string hash) {
        string path = Path.Combine(AppContext.BaseDirectory, "Fonts", name);
        Require(Hash(path) == hash, name + " font changed.");
        return File.ReadAllBytes(path);
    }

    internal static string Hash(string path) => Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path)));

    internal static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
