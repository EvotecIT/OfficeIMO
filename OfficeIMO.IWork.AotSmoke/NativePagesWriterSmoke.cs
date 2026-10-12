using System.Globalization;
using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;

internal static class NativePagesWriterSmoke {
    private const string AcceptedSha256 = "C6822FD3C93E1D9D87CE11270573AEC07AC5AF5167AEFDA1C5C25A8FA683AD53";

    internal static void Run() {
        var document = IWorkPagesDocument.Create();
        document.AddParagraph("OfficeIMO native body document", fontSizePoints: 24);
        document.AddRichParagraph(new[] {
            new IWorkTextRunData("Bold 😀", "Arial", 16, "003366", bold: true, italic: true, underline: true, strikethrough: true),
            new IWorkTextRunData(" plain reset", "Arial", 12),
            new IWorkTextRunData(" — Zażółć café", "Times New Roman", 14, "008000", italic: true)
        });
        document.AddParagraph("");
        for (int index = 1; index <= 80; index++)
            document.AddParagraph("Paragraph " + index.ToString("00", CultureInfo.InvariantCulture) + ": the body flows over pages with editable text.");
        byte[] bytes = document.SaveBytes();
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == AcceptedSha256 && bytes.Length == 47_534,
            "The NativeAOT Pages creator differs from the independently decoded Apple-accepted package.");
        Require(bytes.SequenceEqual(document.SaveBytes()), "Pages creation is not deterministic.");
        using var destination = new MemoryStream(); document.Save(destination);
        Require(destination.CanWrite && bytes.SequenceEqual(destination.ToArray()), "Pages stream save differs or closes its caller.");
        IWorkPagesProjection projection = IWorkSourceDocument.Open(new MemoryStream(bytes)).ReadPages();
        Require(projection.Paragraphs.Count == 84 && projection.Paragraphs[2] == ""
            && projection.Paragraphs[82] == "Paragraph 80: the body flows over pages with editable text."
            && projection.SourceDeclarationIssues.Count == 0 && projection.SourceReferenceIssues.Count == 0,
            "Pages creation lost complete body text or graph declarations.");
        var runs = projection.Body!.Paragraphs[1].Runs;
        Require(runs[0].Text == "Bold 😀" && runs[0].Style is { Bold: true, Italic: true, Underline: true, Strikethrough: true }
            && runs[1].Style is { Bold: false, Italic: false, Underline: false, Strikethrough: false }
            && runs[1].Style!.FontSizePoints == 12 && runs[2].Style!.FontName == "Times New Roman",
            "Pages Unicode ranges or formatting resets differ.");
        using var converted = WordIWorkConverter.ConvertPagesToWordResult(new MemoryStream(bytes),
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        converted.RequireCompleteEditableReconstruction();
        using var saved = new MemoryStream(); converted.Value.Save(saved); saved.Position = 0;
        using var reopened = WordDocument.Load(saved);
        Require(string.Concat(reopened.Paragraphs.Select(paragraph => paragraph.Text)) == string.Concat(projection.Paragraphs),
            "Saved Pages DOCX loses Unicode, body content or text order.");
        try {
            document.Save(destination, new() { MaximumParagraphs = 82 });
            throw new InvalidOperationException("Pages paragraph limit was not enforced.");
        } catch (InvalidDataException) {
            Require(bytes.SequenceEqual(destination.ToArray()), "Pages budget rejection mutated its destination.");
        }
        Console.WriteLine($"PASS | native Pages body creation | {AcceptedSha256} | {bytes.Length} bytes | deterministic graph, rich Unicode paragraphs, plain resets, page settings, caller stream, strict saved/reopened DOCX and paragraph bounds");
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
