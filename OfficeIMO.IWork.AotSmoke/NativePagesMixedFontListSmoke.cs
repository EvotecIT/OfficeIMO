using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;
using W = DocumentFormat.OpenXml.Wordprocessing;

internal static class NativePagesMixedFontListSmoke {
    internal static void Run() {
        var document = Example(); byte[] bytes = document.SaveBytes();
        Require(bytes.Length == 20753 && Hash(bytes) == "92D388CDE48F4B8DB891E39F25530E19A79FB63343A081BDD2A336BFE80AFB07"
            && bytes.SequenceEqual(document.SaveBytes()), "Mixed-family creation differs from the Apple-accepted package.");
        Verify(bytes, edited: false);
        byte[] original = Read("original");
        Require(Hash(original) == "7E9E9A9C7454619ADE26C4F8263156C7D84136677D26D5AA442128128CCD4053", "Apple mixed-family save changed.");
        Verify(original, edited: false);
        byte[] edited = Read("edited");
        Require(Hash(edited) == "9A566E15F005B20D018B0771FF6AA5459891A286EB3CC71498A8B554737EEF90", "Apple mixed-family edit changed.");
        Verify(edited, edited: true);
        var list = document.NumberedLists[0];
        try {
            list.AddRichItem(new[] { new IWorkTextRunData("Fractional size", "Courier New", 14.5f) }, restartAt: 100);
            throw new InvalidOperationException("Mixed-family input bypassed the whole-point size restriction.");
        } catch (ArgumentException) { }
        Require(bytes.SequenceEqual(document.SaveBytes()) && list.AddItem("Next child", level: 1).ListNumber == 3,
            "Rejected fractional-size input changed the package or list counters.");
        using var output = new MemoryStream(); output.WriteByte(9);
        try {
            document.Save(output, new() { MaximumTextRuns = 1 });
            throw new InvalidOperationException("Mixed-family runs bypassed the save budget.");
        } catch (InvalidDataException) { }
        Require(output.ToArray().SequenceEqual(new byte[] { 9 }), "Run-budget rejection changed the caller stream.");
        Console.WriteLine("PASS | native Pages mixed font families | deterministic creation, nine levels, independent marker fonts, Apple edit/insertion/reopen, strict saved DOCX fonts and atomic size/run bounds");
    }

    internal static void Verify(byte[] bytes, bool edited, Func<string, double[]>? runSizes = null) {
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
        Require(projection.Body.IsFormattingComplete && projection.Diagnostics.Count == 0
            && projection.SourceDeclarationIssues.Count == 0 && projection.SourceReferenceIssues.Count == 0,
            "Mixed-family source contains unqualified declarations.");
        var items = projection.Body.Paragraphs.Where(item => item.ListNumber.HasValue).ToArray();
        int[] numbers = [7, 1, 1, 1, 1, 1, 1, 1, 1, 2, 8, 1, 3, 1, 5, 1, 2];
        int[] levels = [0, 1, 2, 3, 4, 5, 6, 7, 8, 8, 0, 1, 0, 1, 0, 1, 1];
        Require(items.Select(item => item.ListNumber!.Value).SequenceEqual(edited ? numbers.Concat(new[] { 3 }) : numbers)
            && items.Select(item => item.ListLevel).SequenceEqual(edited ? levels.Concat(new[] { 1 }) : levels),
            "Mixed-family counters, levels or native insertion changed.");
        Require(items[0].Text == (edited ? "Root native" : "Root seven") + " serif mono plain"
            && items[8].Text == (edited ? "Deep native" : "Depth 8") + " serif mono plain"
            && (!edited || items[^1].Text == "Native child plain"), "Native mixed-family text edits did not survive reopen.");
        foreach (var item in items) {
            string marker = item.Text.StartsWith("Serif ", StringComparison.Ordinal) ? "Times New Roman"
                : item.Text.StartsWith("Mono ", StringComparison.Ordinal) ? "Courier New" : "Arial";
            Require(item.ListFontName == marker && item.Style.TextStyle.FontName == marker
                && item.Style.TextStyle.FontSizePoints == 12 && item.Style.TextStyle.Bold == false
                && item.Style.TextStyle.Italic == false && item.Style.TextStyle.Color?.RgbHex == "000000",
                "Character font or emphasis leaked into the marker.");
            if (item.Text == "Native child plain") {
                Require(item.Runs.Count == 1 && item.Runs[0].Style.FontName == "Arial"
                    && item.Runs[0].Style.FontSizePoints == (runSizes?.Invoke(item.Text)[0] ?? 12),
                    "Native insertion lost its font or size.");
                continue;
            }
            Require(item.Runs.Select(run => run.Style.FontName).SequenceEqual(new[] { "Times New Roman", "Courier New", "Arial" })
                && item.Runs.Select(run => run.Style.FontSizePoints!.Value)
                    .SequenceEqual(runSizes?.Invoke(item.Text) ?? new double[] { 12, 12, 12 })
                && item.Runs[0].Style.Bold == true && item.Runs[0].Style.Color?.RgbHex == "003366"
                && item.Runs[1].Style.Italic == true && item.Runs[1].Style.Color?.RgbHex == "008000"
                && item.Runs[2].Style.Bold == false && item.Runs[2].Style.Italic == false,
                "Mixed-family character runs or reset changed.");
        }
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction(); using var output = new MemoryStream(); result.Value.Save(output);
        using var reopened = WordDocument.Load(new MemoryStream(output.ToArray()));
        Require(!reopened.ValidateDocument().Any() && WordDocumentTraversal.BuildListIndices(reopened).Values
            .Select(item => item.Index).SequenceEqual(items.Select(item => item.ListNumber!.Value)), "Saved DOCX lost counters or validity.");
        output.Position = 0; using var package = WordprocessingDocument.Open(output, false);
        var numbered = package.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>()
            .Where(item => item.ParagraphProperties?.NumberingProperties != null).ToArray();
        Require(numbered.Length == items.Length, "Saved DOCX lost numbered paragraphs.");
        for (int index = 0; index < items.Length; index++) {
            Require(numbered[index].Elements<W.Run>().Where(run => run.InnerText.Length > 0)
                .Select(run => run.RunProperties!.RunFonts!.Ascii!.Value).SequenceEqual(items[index].Runs.Select(run => run.Style.FontName)),
                "Saved DOCX lost explicit run font families.");
            int numberId = numbered[index].ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
            var numbering = package.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
            int abstractId = numbering.Elements<W.NumberingInstance>().Single(value => value.NumberID!.Value == numberId)
                .AbstractNumId!.Val!.Value;
            var level = numbering.Elements<W.AbstractNum>().Single(value => value.AbstractNumberId!.Value == abstractId)
                .Elements<W.Level>().Single(value => value.LevelIndex!.Value == items[index].ListLevel);
            Require(level.NumberingSymbolRunProperties!.FontSize!.Val!.Value
                == (items[index].Runs[0].Style.FontSizePoints!.Value * 2).ToString(System.Globalization.CultureInfo.InvariantCulture),
                "Saved DOCX marker size differs from the first character run.");
            Require(numbered[index].ParagraphProperties!.ParagraphMarkRunProperties!.GetFirstChild<W.FontSize>()!.Val!.Value
                == level.NumberingSymbolRunProperties.FontSize.Val.Value,
                "Paragraph-mark size would rescale markers on Pages import.");
        }
    }

    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Mixed family decimal items");
        var primary = document.CreateNumberedList("Arial", 12);
        primary.AddRichItem(Runs("Root seven"), restartAt: 7);
        for (int level = 1; level <= 8; level++) primary.AddRichItem(Runs("Depth " + level), level: level);
        primary.AddRichItem(Runs("Deep two"), level: 8); document.AddParagraph("Plain gap retains counters");
        primary.AddRichItem(Runs("Root eight")); primary.AddRichItem(Runs("Child one"), level: 1);
        var serif = document.CreateNumberedList(IWorkDecimalListMarkerStyle.Parentheses, "Times New Roman", 12);
        serif.AddRichItem(Runs("Serif marker"), restartAt: 3); serif.AddRichItem(Runs("Serif child"), level: 1);
        var mono = document.CreateNumberedList(IWorkDecimalListMarkerStyle.ClosingParenthesis, "Courier New", 12);
        mono.AddRichItem(Runs("Mono marker"), restartAt: 5); mono.AddRichItem(Runs("Mono child"), level: 1);
        primary.AddRichItem(Runs("Resumed child"), level: 1); return document;
    }

    private static IWorkTextRunData[] Runs(string text) => new[] {
        new IWorkTextRunData(text + " serif", "Times New Roman", 12, "003366", bold: true),
        new IWorkTextRunData(" mono", "Courier New", 12, "008000", italic: true), new IWorkTextRunData(" plain", "Arial", 12)
    };
    private static byte[] Read(string suffix) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "pages-mixed-font-" + suffix + ".pages"));
    private static string Hash(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));
    private static void Require(bool condition, string message) { if (!condition) throw new InvalidOperationException(message); }
}
