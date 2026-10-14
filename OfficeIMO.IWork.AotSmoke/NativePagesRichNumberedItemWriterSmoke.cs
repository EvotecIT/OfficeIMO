using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;
using W = DocumentFormat.OpenXml.Wordprocessing;

internal static class NativePagesRichNumberedItemWriterSmoke {
    internal static void Run() {
        var document = Example(); byte[] bytes = document.SaveBytes();
        Require(Convert.ToHexString(SHA256.HashData(bytes)) == "9AAB5AA4C78F2F0A427255D8126C6CA80D44CCFFB9E23E51854E10E862AC898A"
            && bytes.SequenceEqual(document.SaveBytes()), "Rich native Pages creation differs from the Apple-accepted package.");
        Verify(bytes, "Resumed parent ten");
        byte[] edited = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "pages-created-rich-numbering-edited.pages"));
        Require(Convert.ToHexString(SHA256.HashData(edited)) == "817ABBEE08E66209C4282EC04D2A2DCF35F9CC2F4CA0D89F8D1B2B49ADEDB880",
            "Apple-edited rich Pages fixture changed.");
        Verify(edited, "Resumed parent ten native edit");
        var list = document.NumberedLists[0];
        try {
            list.AddRichItem(new[] { new IWorkTextRunData("Fractional size", "Arial", 14.5f) }, restartAt: 100);
            throw new InvalidOperationException("Fractional rich-list size was accepted.");
        } catch (ArgumentException) { }
        Require(bytes.SequenceEqual(document.SaveBytes()) && list.AddItem("Next").ListNumber == 11,
            "Rejected rich-list input changed counters or document state.");
        using var output = new MemoryStream(); output.WriteByte(9);
        try {
            document.Save(output, new() { MaximumTextRuns = 1 });
            throw new InvalidOperationException("Rich-list run budget was not enforced.");
        } catch (InvalidDataException) { }
        Require(output.ToArray().SequenceEqual(new byte[] { 9 }), "Run-budget rejection changed the caller stream.");
        Console.WriteLine("PASS | native Pages rich decimal-list creation | deterministic native bytes, Apple edit/save/reopen, independent marker/body styles, strict saved DOCX and atomic bounds");
    }

    private static void Verify(byte[] bytes, string lastText) {
        var source = IWorkSourceDocument.Open(bytes); var projection = source.ReadPages();
        Require(projection.Diagnostics.Count == 0 && projection.SourceDeclarationIssues.Count == 0
            && projection.SourceReferenceIssues.Count == 0 && projection.Body.IsFormattingComplete,
            "Rich Pages source has unqualified declarations.");
        var items = projection.Body.Paragraphs.Where(paragraph => paragraph.ListNumber.HasValue).ToArray();
        int[] counters = [7, 1, 5, 8, 1, 9, 3, 10];
        Require(items.Select(item => item.ListNumber!.Value).SequenceEqual(counters) && items[^1].Text == lastText,
            "Rich Pages counters or native edit changed.");
        Require(items.All(item => item.Style.TextStyle.Bold == false && item.Style.TextStyle.Italic == false
            && item.Style.TextStyle.Underline == false && item.Style.TextStyle.Strikethrough == false
            && item.Style.TextStyle.Color?.RgbHex == "000000"), "First-run emphasis leaked into list markers.");
        Require(items[0].Runs[0].Style.Bold == true && items[0].Runs[0].Style.Color?.RgbHex == "003366"
            && items[0].Runs[1].Style.Bold == false && items[0].Runs[1].Style.Color?.RgbHex == "000000"
            && items[1].Runs[0].Style.Italic == true && items[1].Runs[1].Style.Underline == true
            && items[1].Runs[2].Style.Strikethrough == true
            && items[2].Runs[0].Style.Bold == true && items[2].Runs[0].Style.Italic == true
            && items[2].Runs[0].Style.Underline == true && items[2].Runs[0].Style.Strikethrough == true,
            "Rich Pages character decorations or resets changed.");
        using var result = source.ToWordDocumentResult(new() { Mode = IWorkConversionMode.EditableOnly });
        result.RequireCompleteEditableReconstruction(); using var output = new MemoryStream(); result.Value.Save(output); output.Position = 0;
        using var saved = WordDocument.Load(new MemoryStream(output.ToArray()));
        Require(!saved.ValidateDocument().Any()
            && WordDocumentTraversal.BuildListIndices(saved).Values.Select(item => item.Index).SequenceEqual(counters),
            "Saved rich-list DOCX lost counters or validity.");
        using var package = WordprocessingDocument.Open(output, false);
        Require(package.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().Any(paragraph => paragraph.InnerText == lastText),
            "Saved rich-list DOCX lost rich paragraph content.");
    }

    private static IWorkPagesDocument Example() {
        var document = IWorkPagesDocument.Create(); document.AddParagraph("Rich numbered items 😀");
        var list = document.CreateNumberedList();
        list.AddRichItem(new[] { new IWorkTextRunData("Bold first 😀", "Arial", 12, "003366", bold: true),
            new IWorkTextRunData(" plain reset", "Arial", 12) }, restartAt: 7);
        list.AddRichItem(new[] { new IWorkTextRunData("Child italic", "Arial", 12, "008000", italic: true),
            new IWorkTextRunData(", underline", "Arial", 12, "003366", underline: true),
            new IWorkTextRunData(", strike", "Arial", 12, strikethrough: true) }, level: 1);
        list.AddRichItem(new[] { new IWorkTextRunData("Combined decoration", "Arial", 12, "660066", bold: true,
            italic: true, underline: true, strikethrough: true), new IWorkTextRunData(" reset", "Arial", 12) }, level: 1, restartAt: 5);
        document.AddParagraph("Plain gap");
        list.AddRichItem(new[] { new IWorkTextRunData("Colored first", "Arial", 12, "CC0000"),
            new IWorkTextRunData(" parent eight", "Arial", 12) });
        list.AddRichItem(new[] { new IWorkTextRunData("Underlined child", "Arial", 12, underline: true),
            new IWorkTextRunData(" reset", "Arial", 12) }, level: 1);
        list.AddItem("Plain parent nine");
        document.CreateNumberedList().AddRichItem(new[] { new IWorkTextRunData("Separate Zażółć café", "Arial", 12, italic: true) }, restartAt: 3);
        list.AddRichItem(new[] { new IWorkTextRunData("Resumed", "Arial", 12, underline: true),
            new IWorkTextRunData(" parent ten", "Arial", 12) });
        return document;
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
