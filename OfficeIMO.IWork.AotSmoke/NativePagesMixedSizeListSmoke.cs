using System.IO.Compression;
using System.Security.Cryptography;
using OfficeIMO.IWork;

internal static class NativePagesMixedSizeListSmoke {
    internal static void Run() {
        var document = Example(); byte[] bytes = document.SaveBytes();
        Require(bytes.Length == 20751 && Hash(bytes) == "D423D6688E3B6EDBE6A1AD0A0B34BABE257232EAD2ED3D2A5154AE8679D98DC2"
            && bytes.SequenceEqual(document.SaveBytes()), "Mixed-size creation differs from the Apple-accepted package.");
        NativePagesMixedFontListSmoke.Verify(bytes, false, Sizes);
        foreach (var (name, hash, edited) in new[] {
            ("oracle", "790E81E27F3D4FB35CD24955D25F4E886B85DA8D767B22E0EAE4B120530D0635", false),
            ("original", "CE5A910F44C5FEA30980DBE5F3D6EEDD71CB2EE3C52F6B76BCA41A501A3174BA", false),
            ("edited", "6511DF28A9EB82C1A3F437165395F4541D07AFECDCF96E1B8256F3793E15F9D8", true) }) {
            byte[] native = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "pages-mixed-size-" + name + ".pages"));
            Require(Hash(native) == hash, "Apple mixed-size fixture changed: " + name);
            NativePagesMixedFontListSmoke.Verify(native, edited, Sizes);
        }
        var identities = new List<string>();
        foreach (float defaultSize in new[] { 12f, 14f }) {
            var model = IWorkPagesDocument.Create();
            model.CreateNumberedList(fontSizePoints: defaultSize).AddRichItem(Runs("Same body", 18, 8, 14));
            byte[] encoded = model.SaveBytes();
            Require(IWorkSourceDocument.Open(encoded).ReadPages().Body.Paragraphs[0].Style.TextStyle.FontSizePoints == defaultSize,
                "Mixed-size paragraph default changed.");
            using var package = new ZipArchive(new MemoryStream(encoded));
            using var identifier = new StreamReader(package.GetEntry("Metadata/DocumentIdentifier")!.Open());
            identities.Add(identifier.ReadToEnd());
        }
        Require(identities.Distinct(StringComparer.Ordinal).Count() == 2, "Independent paragraph defaults collided in document identity.");
        var list = document.NumberedLists[0];
        try {
            list.AddRichItem(new[] { new IWorkTextRunData("Fractional", "Arial", 14.5f) }, restartAt: 100);
            throw new InvalidOperationException("Fractional numbered run bypassed the qualification boundary.");
        } catch (ArgumentException) { }
        Require(bytes.SequenceEqual(document.SaveBytes()) && list.AddItem("Next child", level: 1).ListNumber == 3,
            "Fractional-size rejection changed native bytes or counters.");
        Console.WriteLine("PASS | native Pages mixed whole-point sizes | first-run markers, default em indents, deterministic identity, Apple edit/insertion/reopen and strict DOCX paragraph marks");
    }

    private static double[] Sizes(string text) => text switch {
        "Root seven serif mono plain" or "Root native serif mono plain" => [18, 8, 14],
        "Root eight serif mono plain" => [8, 18, 14],
        "Child one serif mono plain" => [8, 10, 14],
        "Resumed child serif mono plain" => [12, 24, 8],
        "Native child plain" => [8],
        _ => [12, 12, 12]
    };

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("Mixed size decimal items");
        var primary = document.CreateNumberedList("Arial", 12);
        primary.AddRichItem(Runs("Root seven", 18, 8, 14), restartAt: 7);
        for (int level = 1; level <= 8; level++) primary.AddRichItem(Runs("Depth " + level), level: level);
        primary.AddRichItem(Runs("Deep two"), level: 8);
        document.AddParagraph("Plain gap retains counters");
        primary.AddRichItem(Runs("Root eight", 8, 18, 14));
        primary.AddRichItem(Runs("Child one", 8, 10, 14), level: 1);
        var serif = document.CreateNumberedList(IWorkDecimalListMarkerStyle.Parentheses, "Times New Roman", 12);
        serif.AddRichItem(Runs("Serif marker"), restartAt: 3);
        serif.AddRichItem(Runs("Serif child"), level: 1);
        var mono = document.CreateNumberedList(IWorkDecimalListMarkerStyle.ClosingParenthesis, "Courier New", 12);
        mono.AddRichItem(Runs("Mono marker"), restartAt: 5);
        mono.AddRichItem(Runs("Mono child"), level: 1);
        primary.AddRichItem(Runs("Resumed child", 12, 24, 8), level: 1);
        return document;
    }

    private static IWorkTextRunData[] Runs(string text, float first = 12, float middle = 12, float last = 12) => new[] {
        new IWorkTextRunData(text + " serif", "Times New Roman", first, "003366", bold: true),
        new IWorkTextRunData(" mono", "Courier New", middle, "008000", italic: true),
        new IWorkTextRunData(" plain", "Arial", last)
    };
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
