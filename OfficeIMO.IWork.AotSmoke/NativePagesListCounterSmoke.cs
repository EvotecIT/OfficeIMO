using System.Security.Cryptography;
using OfficeIMO.IWork;
using OfficeIMO.Word.IWork;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;

internal static class NativePagesListCounterSmoke {
    internal static void Run() {
        string path = Path.Combine(AppContext.BaseDirectory, "pages-list-continuation.pages");
        const string Hash = "4696FF27D21DB00E93D7A7985ABA09C22F5F483BF714CE97C680E1A5528A4D53";
        Require(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) == Hash, "Native list fixture changed.");
        using var result = WordIWorkConverter.ConvertPagesToWordResult(path,
            conversionOptions: new() { Mode = IWorkConversionMode.EditableOnly });
        result.Report.RequireCompleteEditableReconstruction();
        var source = result.Projection.Body.Paragraphs;
        Require(result.Projection.Body.IsFormattingComplete
            && source.Select(paragraph => paragraph.ListNumber).SequenceEqual(new int?[] {
                null, 7, 8, null, 9, 1, 2, 10, 1, 5, 6, null, 7, 11, 3, 4 }),
            "Native decimal levels, parent resets or resumed counters differ.");
        using var saved = new MemoryStream(); result.Value.Save(saved); saved.Position = 0;
        using var document = WordprocessingDocument.Open(saved, false);
        var paragraphs = document.MainDocumentPart!.Document!.Body!.Elements<W.Paragraph>().ToArray();
        Require(paragraphs.Select(paragraph => paragraph.InnerText).SequenceEqual(source.Select(paragraph => paragraph.Text)),
            "Saved DOCX changed native list text or paragraph order.");
        var numbering = document.MainDocumentPart.NumberingDefinitionsPart!.Numbering!;
        for (int index = 0; index < paragraphs.Length; index++) {
            if (source[index].ListLevel < 0) continue;
            var properties = paragraphs[index].ParagraphProperties!;
            int level = source[index].ListLevel;
            Require(properties.NumberingProperties!.NumberingLevelReference!.Val!.Value == level
                && properties.Indentation!.Left!.Value == (level == 0 ? "330" : "690")
                && properties.Indentation.Hanging!.Value == "330", "DOCX list level or hanging geometry differs.");
            int id = properties.NumberingProperties.NumberingId!.Val!.Value;
            int abstractId = numbering.Elements<W.NumberingInstance>().Single(number => number.NumberID!.Value == id)
                .AbstractNumId!.Val!.Value;
            W.Level marker = numbering.Elements<W.AbstractNum>().Single(number => number.AbstractNumberId!.Value == abstractId)
                .Elements<W.Level>().Single(item => item.LevelIndex!.Value == level);
            var font = marker.GetFirstChild<W.NumberingSymbolRunProperties>()!;
            Require(font.GetFirstChild<W.FontSize>()!.Val!.Value == "22"
                && font.GetFirstChild<W.RunFonts>()!.Ascii!.Value == "Calibri", "DOCX changed native marker font declarations.");
        }
        Require(Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))) == Hash, "List conversion changed source bytes.");
        Console.WriteLine("PASS | native Pages decimal levels and continuation | " + Hash
            + " | parent resets, explicit restarts, plain-text gaps, strict DOCX save/reopen, marker declarations and geometry");
    }

    private static void Require(bool condition, string message) {
        if (!condition) throw new InvalidOperationException(message);
    }
}
