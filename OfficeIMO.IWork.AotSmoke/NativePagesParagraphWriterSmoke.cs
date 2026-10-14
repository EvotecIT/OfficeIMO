using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesParagraphWriterSmoke {
    internal const string AcceptedSha256 = "689B9A9A31704D48E8092E189F46928F11B30F783CF87F92D7220EDF33FBEBFD";
    internal static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("Native paragraph layout", fontSizePoints: 18);
        document.AddParagraph("Relative spacing 1.5: this paragraph has six points before and nine points after. Its text wraps across lines so the native line spacing can be inspected alongside the exported paragraph settings.",
            new IWorkPagesParagraphLayout(1.5f, 6, 9));
        document.AddParagraph("Start\tLeft\tCenter\tRight\t123.45", new IWorkPagesParagraphLayout(tabStops: new[] {
            IWorkTabStop.Create(90), IWorkTabStop.Create(180, IWorkTabAlignment.Center),
            IWorkTabStop.Create(270, IWorkTabAlignment.Right), IWorkTabStop.Create(360, IWorkTabAlignment.Decimal)
        }));
        document.AddParagraph("Double spacing: this paragraph has twelve points before and eighteen points after. Wrapped lines make the relative spacing visible without changing the font or body width.",
            new IWorkPagesParagraphLayout(2, 12, 18));
        document.AddParagraph("", new IWorkPagesParagraphLayout(1.5f, 6, 6), fontSizePoints: 18);
        document.AddRichParagraph(new[] {
            new IWorkTextRunData("Rich layout", fontSizePoints: 14, bold: true),
            new IWorkTextRunData(" plain Zażółć 😀", fontSizePoints: 12)
        }, new IWorkPagesParagraphLayout(1.25f, spaceAfterPoints: 6));
        document.AddParagraph("Default reset paragraph.");
        return document;
    }

    internal static void Run() {
        var document = Example(); byte[] bytes = document.SaveBytes();
        Require(bytes.Length == 8450 && Convert.ToHexString(SHA256.HashData(bytes)) == AcceptedSha256,
            "Native Pages paragraph layout differs from the Apple-qualified encoding.");
        Require(bytes.SequenceEqual(Example().SaveBytes()), "Pages paragraph layout encoding is not deterministic.");
        var projection = IWorkSourceDocument.Open(new MemoryStream(bytes)).ReadPages();
        Require(projection.SourceDeclarationIssues.Count == 0 && projection.SourceReferenceIssues.Count == 0
            && projection.Diagnostics.Count == 0 && projection.Paragraphs.Count == 8,
            "Native Pages paragraph layout has incomplete source evidence.");
        var paragraphs = projection.Body.Paragraphs;
        Require(paragraphs[1].Style.LineSpacingMultiplier == 1.5 && paragraphs[1].Style.SpaceBeforePoints == 6
            && paragraphs[1].Style.SpaceAfterPoints == 9 && paragraphs[2].Style.TabStops!.Count == 4
            && paragraphs[2].Style.DecimalTabSymbol == "."
            && paragraphs[2].Text == "Start\tLeft\tCenter\tRight\t123.45" && paragraphs[3].Style.LineSpacingMultiplier == 2
            && paragraphs[4].Style.TextStyle.FontSizePoints == 18 && paragraphs[5].Style.LineSpacingMultiplier == 1.25
            && paragraphs[6].Style.LineSpacingMultiplier == 1 && paragraphs[6].Style.TabStops!.Count == 0,
            "Native Pages spacing, tabs or paragraph resets differ.");
        using var converted = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var output = new MemoryStream(); converted.Value.Save(output); output.Position = 0;
        using var reopened = WordDocument.Load(output);
        Require(reopened.Settings.DecimalSymbol == "."
            && string.Concat(reopened.Paragraphs.Select(p => p.Text)) == string.Concat(projection.Paragraphs)
            && reopened.Paragraphs.Any(p => p.Text.StartsWith("Relative spacing 1.5:", StringComparison.Ordinal) && p.LineSpacing == 360)
            && reopened.Paragraphs.Any(p => p.Text.StartsWith("Double spacing:", StringComparison.Ordinal) && p.LineSpacing == 480)
            && reopened.Paragraphs.Any(p => p.Text == "Rich layout" && p.LineSpacing == 300)
            && reopened.Paragraphs.Any(p => p.Text == "Start\tLeft\tCenter\tRight\t123.45" && p.TabStops.Count == 4),
            "Saved/reopened Pages DOCX lost relative line spacing.");
        using var destination = new MemoryStream(); destination.WriteByte(9);
        try {
            document.Save(destination, new() { MaximumTabStops = 3 });
            throw new InvalidOperationException("Pages custom-tab limit was not enforced.");
        } catch (InvalidDataException) {
            Require(destination.ToArray().SequenceEqual(new byte[] { 9 }), "Pages tab-limit rejection changed its destination.");
        }
        Console.WriteLine($"PASS | native Pages paragraph layout | {Convert.ToHexString(SHA256.HashData(bytes))} | {bytes.Length} bytes | spacing, four tab alignments, TAB text, blank metrics, resets, strict DOCX and custom-tab bounds");
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
