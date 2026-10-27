using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using System.Text.RegularExpressions;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(false, "default")]
    [InlineData(false, "first")]
    [InlineData(false, "even")]
    [InlineData(true, "default")]
    [InlineData(true, "first")]
    [InlineData(true, "even")]
    public void SaveAsPdf_InheritedRunningTablePaintsEachImageOnce(bool footer, string kind) {
        using WordDocument source = CreateInheritedRunningStoryControl(footer, kind, "next");
        WordSection first = source.Sections[0];
        WordHeaderFooter story = footer
            ? kind == "first" ? first.Footer.First! : kind == "even" ? first.Footer.Even! : first.Footer.Default!
            : kind == "first" ? first.Header.First! : kind == "even" ? first.Header.Even! : first.Header.Default!;
        foreach (var child in story.ChildElements.ToArray()) child.Remove();
        WordTable table = CreateRunningStoryTable(story, "MARK", 48, WordTextDirection.LeftToRightTopToBottom);
        table.Rows[0].Cells[0].Paragraphs[0].AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), 32, 32);
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        byte[] bytes = document.ToPdf(EmptyRunningOptions()).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(3, pdf.NumberOfPages);
        int expectedPaints = kind == "default" ? 3 : kind == "first" ? 2 : 1;
        Assert.Equal(expectedPaints, pdf.GetPages().Count(page => page.Text.Contains("MARK")));
        // Count content-stream invocations, not deduplicated image resources.
        var (objects, _) = PdfSyntax.ParseObjects(bytes);
        PdfReadDocument read = PdfReadDocument.Open(bytes);
        int paints = read.Pages.Sum(page => {
            var dictionary = Assert.IsType<PdfDictionary>(objects[page.ObjectNumber].Value);
            return CountRunningStoryImagePaints(dictionary.Items["Contents"], objects);
        });
        Assert.Equal(expectedPaints, paints);
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.Document.OuterXml);
    }

    private static int CountRunningStoryImagePaints(PdfObject contents, Dictionary<int, PdfIndirectObject> objects) {
        if (contents is PdfArray array) return array.Items.Sum(item => CountRunningStoryImagePaints(item, objects));
        var reference = Assert.IsType<PdfReference>(contents);
        var stream = Assert.IsType<PdfStream>(objects[reference.ObjectNumber].Value);
        return Regex.Matches(PdfOperatorSearchText.Decode(stream, objects), @"/Im\d+\s+Do\b").Count;
    }
}
