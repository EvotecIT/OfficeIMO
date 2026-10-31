using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    public static IEnumerable<object[]> KeptParagraphContinuationCases() {
        foreach (var source in new[] {
            (WordFileFormat.Docx, WordCompatibilityMode.Word2007, false),
            (WordFileFormat.Docx, WordCompatibilityMode.Word2013, true),
            (WordFileFormat.Doc, WordCompatibilityMode.Word2013, false)
        }) {
            foreach (bool numbered in new[] { false, true }) {
                foreach (bool prelude in new[] { false, true }) {
                    yield return new object[] { source.Item1, source.Item2, source.Item3, numbered, prelude, false, "none", false };
                    if (prelude) yield return new object[] { source.Item1, source.Item2, source.Item3, numbered, prelude, true, "none", false };
                }
            }
            foreach (string decoration in new[] { "shading", "border" }) {
                yield return new object[] { source.Item1, source.Item2, source.Item3, false, false, false, decoration, false };
                yield return new object[] { source.Item1, source.Item2, source.Item3, false, true, false, decoration, false };
                yield return new object[] { source.Item1, source.Item2, source.Item3, false, true, true, decoration, false };
                yield return new object[] { source.Item1, source.Item2, source.Item3, false, true, false, decoration, true };
            }
        }
    }

    [Theory]
    [MemberData(nameof(KeptParagraphContinuationCases))]
    public void SaveAsPdf_OversizedKeptParagraphContinuesUsingSourceColumnCompatibility(
        WordFileFormat format, WordCompatibilityMode mode, bool skipsColumns, bool numbered, bool prelude, bool sectionPrelude,
        string decoration, bool keepWithNext) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.Sections[0].ColumnCount = 2;
        document.Sections[0].ColumnsSpace = global::OfficeIMO.OfficeLength.Points(20);
        if (prelude) {
            WordParagraph before = document.AddParagraph("Before");
            ConfigurePartialColumnParagraph(before);
            before.KeepWithNextOverride = keepWithNext;
        }
        if (sectionPrelude) {
            document.Sections[0].ColumnCount = 1;
            WordSection columns = document.AddSection(WordSectionBreakType.Continuous);
            columns.ColumnCount = 2;
            columns.ColumnsSpace = global::OfficeIMO.OfficeLength.Points(20);
        }
        string text = "Marker001\nMarker002\nMarker003\nMarker004";
        WordParagraph paragraph;
        if (numbered) {
            WordList list = document.AddCustomList();
            list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
            list.Numbering.Levels[0].IndentationLeft= global::OfficeIMO.OfficeLength.Points(12D);
            list.Numbering.Levels[0].IndentationHanging= global::OfficeIMO.OfficeLength.Points(12D);
            paragraph = list.AddItem(text);
        } else {
            paragraph = document.AddParagraph(text);
        }
        ConfigurePartialColumnParagraph(paragraph);
        paragraph.KeepLinesTogetherOverride = true;
        if (decoration == "shading") paragraph.ShadingFillColorHex = "00FFFF";
        if (decoration == "border") {
            paragraph.Borders.TopStyle = paragraph.Borders.BottomStyle =
                paragraph.Borders.LeftStyle = paragraph.Borders.RightStyle = WordBorderStyle.Single;
            paragraph.Borders.TopSize = paragraph.Borders.BottomSize =
                paragraph.Borders.LeftSize = paragraph.Borders.RightSize = OfficeLength.Points(0.5);
            paragraph.Borders.TopSpace = paragraph.Borders.BottomSpace = OfficeLength.Points(0);
        }
        ConfigurePartialColumnParagraph(document.AddParagraph("After"));
        var options = CellPaginationOptions();
        options.PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(500),OfficeIMO.OfficeLength.Points(121));
        options.PdfOptions = new OfficeIMO.Pdf.PdfCreateOptions { MaxGeneratedPages = 6 };
        using var input = new MemoryStream(document.ToBytes(format));
        using WordDocument loaded = WordDocument.Load(input);
        using var output = new MemoryStream();
        loaded.SaveAsPdf(output, options);
        using var pdf = PdfPigDocument.Open(output.ToArray());
        int firstPage = prelude && (skipsColumns || sectionPrelude) ? 2 : 1;
        int secondPage = skipsColumns ? firstPage + 1 : (prelude ? 2 : 1);
        int afterPage = skipsColumns ? secondPage : (sectionPrelude ? 3 : 2);
        Assert.Equal(afterPage, pdf.NumberOfPages);
        double firstMarkerBottom = 0D;
        foreach (int index in Enumerable.Range(1, 4)) {
            int expectedPage = index <= 2 ? firstPage : secondPage;
            var words = pdf.GetPage(expectedPage).GetWords();
            var marker = Assert.Single(words, word => word.Text == $"Marker{index:D3}" ||
                word.Text.EndsWith($".Marker{index:D3}", StringComparison.Ordinal));
            bool rightColumn = !skipsColumns && (index <= 2 ? prelude && !sectionPrelude : !prelude || sectionPrelude);
            double left = rightColumn ? 260 : 40;
            Assert.InRange(marker.BoundingBox.Left, left - .1, left + (numbered ? 40 : .1));
            if (index == 1) firstMarkerBottom = marker.BoundingBox.Bottom;
            if (index == 3 && decoration == "border")
                Assert.Equal(.5D, marker.BoundingBox.Bottom - firstMarkerBottom, 2);
        }
        var allWords = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).ToArray();
        foreach (int index in Enumerable.Range(1, 4)) Assert.Single(allWords, word => word.Text == $"Marker{index:D3}" ||
            word.Text.EndsWith($".Marker{index:D3}", StringComparison.Ordinal));
        var after = Assert.Single(pdf.GetPage(afterPage).GetWords(), word => word.Text == "After");
        double afterX = skipsColumns || prelude && !sectionPrelude ? 260 : 40;
        Assert.InRange(after.BoundingBox.Left, afterX - .1, afterX + .1);
        if (prelude) Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "Before");
    }

    [Theory]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2007)]
    [InlineData(WordFileFormat.Docx, WordCompatibilityMode.Word2013)]
    [InlineData(WordFileFormat.Doc, WordCompatibilityMode.Word2013)]
    public void SaveAsPdf_KeptNumberedParagraphsMoveIndividuallyInsteadOfKeepingTheWholeList(
        WordFileFormat format, WordCompatibilityMode mode) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.Sections[0].ColumnCount = 2;
        document.Sections[0].ColumnsSpace = global::OfficeIMO.OfficeLength.Points(20);
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.Levels[0].IndentationLeft= global::OfficeIMO.OfficeLength.Points(12D);
        list.Numbering.Levels[0].IndentationHanging= global::OfficeIMO.OfficeLength.Points(12D);
        foreach (int item in Enumerable.Range(0, 2)) {
            WordParagraph paragraph = list.AddItem($"Marker{item * 2 + 1:D3}\nMarker{item * 2 + 2:D3}");
            ConfigurePartialColumnParagraph(paragraph);
            paragraph.KeepLinesTogetherOverride = true;
        }
        ConfigurePartialColumnParagraph(document.AddParagraph("After"));
        var options = CellPaginationOptions();
        options.PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(500),OfficeIMO.OfficeLength.Points(121));
        options.PdfOptions = new OfficeIMO.Pdf.PdfCreateOptions { MaxGeneratedPages = 6 };
        using var input = new MemoryStream(document.ToBytes(format));
        using WordDocument loaded = WordDocument.Load(input);
        using var output = new MemoryStream();
        loaded.SaveAsPdf(output, options);
        using var pdf = PdfPigDocument.Open(output.ToArray());
        Assert.Equal(2, pdf.NumberOfPages);
        foreach (int index in Enumerable.Range(1, 4)) {
            var marker = Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == $"Marker{index:D3}" ||
                word.Text.EndsWith($".Marker{index:D3}", StringComparison.Ordinal));
            double left = index <= 2 ? 40 : 260;
            Assert.InRange(marker.BoundingBox.Left, left - .1, left + 40);
        }
        Assert.Single(pdf.GetPage(2).GetWords(), word => word.Text == "After");
    }

    public static IEnumerable<object[]> OneLineParagraphContinuationCases() {
        foreach (var source in new[] {
            (WordFileFormat.Docx, WordCompatibilityMode.Word2007, false),
            (WordFileFormat.Docx, WordCompatibilityMode.Word2013, true),
            (WordFileFormat.Doc, WordCompatibilityMode.Word2013, false)
        }) {
            foreach (string decoration in new[] { "none", "shading", "border" }) {
                foreach (bool keepTogether in new[] { false, true }) {
                    foreach (bool widowControl in new[] { false, true }) {
                        yield return new object[] { source.Item1, source.Item2,
                            source.Item3 && (keepTogether || widowControl), decoration, keepTogether, widowControl };
                    }
                }
            }
        }
    }

    [Theory]
    [MemberData(nameof(OneLineParagraphContinuationCases))]
    public void SaveAsPdf_OneLineParagraphFramesRelaxUnavoidableWidowsUsingSourceCompatibility(
        WordFileFormat format, WordCompatibilityMode mode, bool skipsColumns, string decoration,
        bool keepTogether, bool widowControl) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        document.Sections[0].ColumnCount = 2;
        document.Sections[0].ColumnsSpace = global::OfficeIMO.OfficeLength.Points(20);
        WordParagraph paragraph = document.AddParagraph("Marker001\nMarker002\nMarker003\nMarker004");
        ConfigurePartialColumnParagraph(paragraph);
        paragraph.KeepLinesTogetherOverride = keepTogether;
        paragraph.AvoidWidowAndOrphanOverride = widowControl;
        if (decoration == "shading") paragraph.ShadingFillColorHex = "00FFFF";
        if (decoration == "border") {
            paragraph.Borders.TopStyle = paragraph.Borders.BottomStyle =
                paragraph.Borders.LeftStyle = paragraph.Borders.RightStyle = WordBorderStyle.Single;
            paragraph.Borders.TopSize = paragraph.Borders.BottomSize =
                paragraph.Borders.LeftSize = paragraph.Borders.RightSize = OfficeLength.Points(0.5);
            paragraph.Borders.TopSpace = paragraph.Borders.BottomSpace = OfficeLength.Points(0);
        }
        ConfigurePartialColumnParagraph(document.AddParagraph("After"));
        var options = CellPaginationOptions();
        options.PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(500),OfficeIMO.OfficeLength.Points(101));
        options.PdfOptions = new OfficeIMO.Pdf.PdfCreateOptions { MaxGeneratedPages = 6 };
        using var input = new MemoryStream(document.ToBytes(format));
        using WordDocument loaded = WordDocument.Load(input);
        using var output = new MemoryStream();
        loaded.SaveAsPdf(output, options);
        using var pdf = PdfPigDocument.Open(output.ToArray());
        Assert.Equal(skipsColumns ? 4 : 3, pdf.NumberOfPages);
        foreach (int index in Enumerable.Range(1, 4)) {
            int page = skipsColumns ? index : (index + 1) / 2;
            var marker = Assert.Single(pdf.GetPage(page).GetWords(), word => word.Text == $"Marker{index:D3}");
            double left = skipsColumns || index % 2 == 1 ? 40 : 260;
            Assert.InRange(marker.BoundingBox.Left, left - .1, left + .1);
        }
        var after = Assert.Single(pdf.GetPage(skipsColumns ? 4 : 3).GetWords(), word => word.Text == "After");
        Assert.InRange(after.BoundingBox.Left, skipsColumns ? 259.9 : 39.9, skipsColumns ? 260.1 : 40.1);
        Assert.Equal(5, Enumerable.Range(1, pdf.NumberOfPages).Sum(page => pdf.GetPage(page).GetWords().Count()));
    }
}
