using System.IO.Compression;
using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesMarkerIdentitySmoke {
    internal static void Run() {
        foreach (var (mode, periodHash) in new[] {
            ("plain", "B70D7651D395916396F79D638B1E3D65A7E1E775B105E2E7AB369C2A98A7DB77"),
            ("mixed-font", "C31F84AD2C73595132FF336D24CC89921E33513DFD5B0689552F6EC6C7CC04B6"),
            ("mixed-size", "557D306CA88E514FF5A3F25BD17B666CAC01DF55382DB212FCDABFFD98D35C3C") }) {
            var identifiers = new List<string>();
            foreach (var style in Enum.GetValues<IWorkDecimalListMarkerStyle>()) {
                var document = Example(mode, style); byte[] bytes = document.SaveBytes();
                Require(bytes.SequenceEqual(Example(mode, style).SaveBytes()), "Marker identity is nondeterministic.");
                if (style == IWorkDecimalListMarkerStyle.Period)
                    Require(Hash(bytes) == periodHash, "Original period-marker bytes changed.");
                using var package = new ZipArchive(new MemoryStream(bytes));
                using var reader = new StreamReader(package.GetEntry("Metadata/DocumentIdentifier")!.Open());
                identifiers.Add(reader.ReadToEnd());
                document.CreateNumberedList(style == IWorkDecimalListMarkerStyle.Period
                    ? IWorkDecimalListMarkerStyle.Parentheses : IWorkDecimalListMarkerStyle.Period);
                Require(bytes.SequenceEqual(document.SaveBytes()), "An unused group changed document identity.");
            }
            Require(identifiers.Distinct(StringComparer.Ordinal).Count() == 3, "Different marker punctuation shares document identity.");
        }
        foreach (var (name, hash) in new[] {
            ("parentheses", "1208453AE59349455C0660E0A7C1268BDAFA0F989D20BCFCB23F6DC3374B21C1"),
            ("closing", "E148045CA79207CD16BB1E868C8FF557F1853FAF5DC26ACECE8C18CECF3E1333"),
            ("mixed-font", "21D1B6C7090D9E53F0F0AC2FAEB2D343B78B040337E21904B3BA0C9F7B2DACFC"),
            ("mixed-size", "B03000EEBFBE52C1D34E7D034B3B621EF44174E4B7D617585943DD5E68B170D6") }) {
            byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "pages-identity-" + name + ".pages"));
            Require(Hash(bytes) == hash, "Fresh Apple marker-identity fixture changed: " + name);
            IWorkPagesDocument original = name switch {
                "parentheses" => NativePagesNineLevelSmoke.Example(IWorkDecimalListMarkerStyle.Parentheses),
                "closing" => NativePagesNineLevelSmoke.Example(IWorkDecimalListMarkerStyle.ClosingParenthesis),
                "mixed-font" => NativePagesMixedFontListSmoke.Example(),
                _ => NativePagesMixedSizeListSmoke.Example()
            };
            var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
            var before = IWorkSourceDocument.Open(original.SaveBytes()).ReadPages().Body.Paragraphs.Where(p => p.Text.Length > 0).ToArray();
            var after = projection.Body.Paragraphs.Where(p => p.Text.Length > 0).ToArray();
            Require(projection.Body.IsFormattingComplete && projection.Diagnostics.Count == 0
                && projection.SourceDeclarationIssues.Count == 0 && projection.SourceReferenceIssues.Count == 0
                && before.Length == after.Length, "Native identity save lost qualified content.");
            for (int index = 0; index < before.Length; index++) {
                Require(after[index].Text == before[index].Text.Replace("Root seven", "Root native", StringComparison.Ordinal)
                    && after[index].ListNumber == before[index].ListNumber && after[index].ListLevel == before[index].ListLevel
                    && after[index].ListLabel == before[index].ListLabel, "Native identity edit changed text or list semantics.");
                Require(after[index].Runs.Select(p => (p.Style.FontName, p.Style.FontSizePoints))
                    .SequenceEqual(before[index].Runs.Select(p => (p.Style.FontName, p.Style.FontSizePoints))),
                    "Native identity save lost run families or sizes.");
            }
            using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
            result.RequireCompleteEditableReconstruction(); using var output = new MemoryStream(); result.Value.Save(output);
            using var saved = WordDocument.Load(new MemoryStream(output.ToArray()));
            var items = after.Where(p => p.ListNumber.HasValue).ToArray();
            string[] markers = items.Select(p => p.ListLabel == "(1)" ? "(" + p.ListNumber + ")"
                : p.ListNumber + (p.ListLabel == "1)" ? ")" : ".")).ToArray();
            Require(!saved.ValidateDocument().Any() && WordDocumentTraversal.BuildListIndices(saved).Values
                .Select(p => p.Index).SequenceEqual(items.Select(p => p.ListNumber!.Value))
                && WordDocumentTraversal.BuildListMarkers(saved).Values.Select(p => p.Marker).SequenceEqual(markers),
                "Strict saved DOCX lost native identity-fixture counters or markers.");
        }
        Console.WriteLine("PASS | native Pages marker punctuation identity | nine deterministic identity cases, unchanged period bytes, unused groups, four fresh Apple edit/save/reopen fixtures and strict saved DOCX");
    }

    private static IWorkPagesDocument Example(string mode, IWorkDecimalListMarkerStyle style) {
        var document = IWorkPagesDocument.Create(); var list = document.CreateNumberedList(style, "Arial", 12);
        if (mode == "plain") list.AddItem("Same marker body");
        else list.AddRichItem(new[] {
            new IWorkTextRunData("Same marker", "Times New Roman", mode == "mixed-size" ? 18 : 12, bold: true),
            new IWorkTextRunData(" body", "Courier New", mode == "mixed-size" ? 8 : 12, italic: true)
        });
        return document;
    }
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
