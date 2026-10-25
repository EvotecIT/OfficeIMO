using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word.Rtf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfPaginationHeaderNativeRegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExactHeaderPicturePreservesNativeLegacyFullFrame(bool throughWord) {
        // This independently authored RTF exports as one page in native Word.
        // Its inline VML carrier paints the whole 24x12pt image. Controlled
        // DOCX exports isolate the carrier, rather than compatibility mode,
        // as the distinction from DrawingML's exact-line picture crop.
        byte[] bytes = ConvertFixture("r01-header-inline-picture-exact", throughWord);
        using var pdf = PdfPigDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var image = Assert.Single(PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes));
        Assert.Equal(1, image.PageNumber);
        Assert.Equal(24D, image.X, 2);
        Assert.Equal(24D, image.Width, 2);
        Assert.Equal(12D, image.Height, 2);
        Assert.Equal(15.6D, page.Height - image.Y - image.Height, 2);
        var words = page.GetWords().ToArray();
        var inline = Assert.Single(words, word => word.Text == "INLINE");
        var title = Assert.Single(words, word => word.Text == "TITLE");
        Assert.True(inline.Letters[0].StartBaseLine.X >= image.X + image.Width - .05D);
        Assert.True(title.Letters[0].StartBaseLine.X >= inline.Letters.Last().EndBaseLine.X);
        Assert.Equal(inline.Letters[0].StartBaseLine.Y, title.Letters[0].StartBaseLine.Y, 2);
        Assert.Equal(27.6D, page.Height - title.Letters[0].StartBaseLine.Y, 2);
        if (image.Clip is { } clip) {
            Assert.True(clip.IsExact && clip.IsRectangle && !clip.ContainsTextClipping);
            double top = page.Height - image.Y - image.Height;
            Assert.True(clip.X <= image.X + .01D && clip.X + clip.Width >= image.X + image.Width - .01D);
            Assert.True(clip.Y <= top + .01D && clip.Y + clip.Height >= top + image.Height - .01D,
                "The legacy RTF image must retain its full painted frame, including the exact-line overhang.");
        }
        string text = ReadPageText(bytes)[0];
        Assert.Equal(1, Occurrences(text, "R01INTRO"));
        Assert.Equal(1, Occurrences(text, "R01FOOTER"));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FirstEvenAndRestartedSectionMatchNativeSixPages(bool throughWord) {
        // Native Word exports six physical pages, with a truly blank fourth
        // page before the restarted second section. This is an observed source
        // contract, not a general rule inferred from the section break token.
        byte[] bytes = ConvertFixture("r02-first-even-section-nextpage", throughWord);
        string[] pages = ReadPageText(bytes);
        Assert.Equal(6, pages.Length);
        string[][] expected = {
            new[] { "R02S1FIRSTHEADER", "R02S1FIRSTFOOTER", "R02S1PAGEONE" },
            new[] { "R02S1EVENHEADER", "R02S1EVENFOOTER", "R02S1PAGETWO" },
            new[] { "R02S1ODDHEADER", "R02S1ODDFOOTER", "R02S1PAGETHREE" },
            Array.Empty<string>(),
            new[] { "R02S2FIRSTHEADER", "R02S2FIRSTFOOTER", "R02S2PAGEONE" },
            new[] { "R02S2EVENHEADER", "R02S2EVENFOOTER", "R02S2PAGETWO" }
        };
        string[] markers = expected.SelectMany(items => items).ToArray();
        for (int index = 0; index < pages.Length; index++) {
            foreach (string marker in markers) {
                Assert.Equal(expected[index].Contains(marker) ? 1 : 0, Occurrences(pages[index], marker));
            }
        }
        Assert.Empty(pages[3]);
        Assert.DoesNotContain(PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes), image => image.PageNumber == 4);
        foreach (string body in markers.Where(marker => marker.Contains("PAGE"))) {
            Assert.Equal(1, Occurrences(string.Concat(pages), body));
        }
    }

    private static byte[] ConvertFixture(string id, bool throughWord) {
        string path = Path.Combine(AppContext.BaseDirectory, "Documents", "PaginationHeaderControls", id + ".rtf");
        byte[] original = File.ReadAllBytes(path);
        RtfReadResult read = RtfDocument.LoadResult(original);
        string before = read.Document.ToRtf(new RtfWriteOptions { IncludeGenerator = false });
        var parse = new RtfConversionReport();
        parse.AddReadDiagnostics(read.Diagnostics, path);
        byte[]? mappedDocx = null;
        RtfConversionReport? mappedReport = null;
        if (throughWord) {
            RtfConversionResult<WordDocument> mapped = read.ToWordDocumentResult(path);
            using WordDocument word = mapped.Value;
            using var docx = new MemoryStream();
            word.Save(docx);
            mappedDocx = docx.ToArray();
            mappedReport = mapped.Report;
        }
        // PDF running stories are materialized during Save. Keep the reopened
        // Word source alive until serialization and its evidence are complete.
        using WordDocument? reopened = throughWord ? WordDocument.Load(new MemoryStream(mappedDocx!)) : null;
        PdfCore.PdfDocumentConversionResult result = throughWord
            ? reopened!.ToPdfDocumentResult().WithSourceConversionReport(mappedReport!)
            : read.Document.ToPdfDocumentResult().WithSourceConversionReport(parse);
        using var output = new MemoryStream();
        PdfCore.PdfSaveResult saved = result.Save(output);
        byte[] bytes = output.ToArray();
        if (Environment.GetEnvironmentVariable("OFFICEIMO_PDF_VISUAL_OUTPUT") is { Length: > 0 } directory) {
            Directory.CreateDirectory(directory);
            string stem = Path.Combine(directory, id + (throughWord ? ".composed" : ".direct"));
            File.WriteAllBytes(stem + ".pdf", bytes);
            if (mappedDocx != null) File.WriteAllBytes(stem + ".mapped.docx", mappedDocx);
            var report = new StringBuilder()
                .AppendLine("Route=" + (throughWord ? "RTF->Word->PDF" : "RTF->PDF"))
                .AppendLine("SaveSucceeded=" + saved.Succeeded)
                .AppendLine("HasLoss=" + saved.HasLoss)
                .AppendLine("SourceBytesUnchanged=" + original.SequenceEqual(File.ReadAllBytes(path)))
                .AppendLine("LosslessBytesUnchanged=" + original.SequenceEqual(read.ToBytesLossless()))
                .AppendLine("LoadedModelUnchanged=" + (before == read.Document.ToRtf(new RtfWriteOptions { IncludeGenerator = false })));
            foreach (var diagnostic in read.Diagnostics) report.AppendLine("Parse: " + diagnostic);
            foreach (var diagnostic in saved.FidelityDiagnostics) {
                report.AppendLine($"Fidelity: {diagnostic.Source}|{diagnostic.Code}|{diagnostic.LossKind}|{diagnostic.Location}|{diagnostic.Message}");
            }
            foreach (string diagnostic in saved.Diagnostics) report.AppendLine("Save: " + diagnostic);
            if (saved.Succeeded) {
                string[] pages = ReadPageText(bytes);
                report.AppendLine("PageCount=" + pages.Length);
                for (int index = 0; index < pages.Length; index++) report.AppendLine($"Page{index + 1}={pages[index]}");
                foreach (var image in PdfCore.PdfImageExtractor.ExtractImagePlacements(bytes)) {
                    report.AppendLine(FormattableString.Invariant($"Image: page={image.PageNumber} x={image.X} y={image.Y} width={image.Width} height={image.Height}"));
                    if (image.Clip is { } clip) report.AppendLine(FormattableString.Invariant($"Clip: x={clip.X} y={clip.Y} width={clip.Width} height={clip.Height} exact={clip.IsExact} rectangle={clip.IsRectangle}"));
                }
            }
            File.WriteAllText(stem + ".report.txt", report.ToString());
        }
        saved.RequireSuccess();
        Assert.Equal(original, File.ReadAllBytes(path));
        Assert.Equal(original, read.ToBytesLossless());
        Assert.Equal(before, read.Document.ToRtf(new RtfWriteOptions { IncludeGenerator = false }));
        return bytes;
    }

    private static string[] ReadPageText(byte[] bytes) => PdfCore.PdfReadDocument.Open(bytes,
        new PdfCore.PdfLoadOptions { IncludeArtifactText = true }).Pages
        .Select(page => new string(page.ExtractText().Where(character => !char.IsWhiteSpace(character)).ToArray())).ToArray();

    private static int Occurrences(string text, string marker) {
        int count = 0;
        for (int index = 0; (index = text.IndexOf(marker, index, StringComparison.Ordinal)) >= 0; index += marker.Length) count++;
        return count;
    }
}
