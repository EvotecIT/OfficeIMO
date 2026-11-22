using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Letter = UglyToad.PdfPig.Content.Letter;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordPdfRightToLeftAlignmentTests {
    private const string Filler = "alpha beta gamma delta epsilon zeta eta theta iota kappa lambda mu nu xi omicron pi rho sigma tau upsilon phi chi psi";

    // Word places w:jc start/left on the right of a bidi paragraph and end/right on its left; end aligns right in an LTR paragraph.
    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("nestedcell")]
    [InlineData("header")]
    [InlineData("textbox")]
    public void SaveAsPdf_ParagraphAlignmentFollowsTheParagraphDirection(string container) {
        var cases = new (string Word, bool BiDi, W.JustificationValues? Alignment, bool Right)[] {
            ("RTLNONE", true, null, true),
            ("RTLLEFT", true, W.JustificationValues.Left, true),
            ("RTLSTART", true, W.JustificationValues.Start, true),
            ("RTLRIGHT", true, W.JustificationValues.Right, false),
            ("RTLEND", true, W.JustificationValues.End, false),
            ("LTRSTART", false, W.JustificationValues.Start, false),
            ("LTREND", false, W.JustificationValues.End, true)
        };
        var paragraphs = new List<W.Paragraph> {
            CreateParagraph("LEFTREF", false, W.JustificationValues.Left),
            CreateParagraph("RIGHTREF", false, W.JustificationValues.Right)
        };
        paragraphs.AddRange(cases.Select(item => CreateParagraph(item.Word, item.BiDi, item.Alignment)));
        W.Paragraph styled = CreateParagraph("RTLSTYLE", false, W.JustificationValues.Left);
        styled.ParagraphProperties!.ParagraphStyleId = new W.ParagraphStyleId { Val = RightToLeftStyleId };
        paragraphs.Add(styled);

        using var pdf = PdfPigDocument.Open(Convert(container, paragraphs));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToList();
        double left = Word(letters, "LEFTREF").First().StartBaseLine.X;
        double right = Word(letters, "RIGHTREF").Last().EndBaseLine.X;
        foreach (var item in cases) {
            List<Letter> word = Word(letters, item.Word);
            double edge = item.Right ? word.Last().EndBaseLine.X : word.First().StartBaseLine.X;
            double expected = item.Right ? right : left;
            Assert.True(Math.Abs(edge - expected) < .05D, $"{item.Word} sits at {edge:F2}; expected {expected:F2}.");
        }
        // The direction can come from the paragraph style.
        Assert.InRange(Word(letters, "RTLSTYLE").Last().EndBaseLine.X, right - .05D, right + .05D);
    }

    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("nestedcell")]
    [InlineData("header")]
    [InlineData("textbox")]
    public void SaveAsPdf_JustifiedRightToLeftParagraphEndsItsLastLineOnTheRight(string container) {
        var paragraphs = new List<W.Paragraph> {
            CreateParagraph("LEFTREF", false, W.JustificationValues.Left),
            CreateParagraph("RIGHTREF", false, W.JustificationValues.Right),
            CreateParagraph("JRTL " + Filler + " JRTLEND", true, W.JustificationValues.Both),
            CreateParagraph("JLTR " + Filler + " JLTREND", false, W.JustificationValues.Both),
            CreateParagraph("DRTL " + Filler + " DRTLEND", true, W.JustificationValues.Distribute)
        };

        using var pdf = PdfPigDocument.Open(Convert(container, paragraphs));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToList();
        double left = Word(letters, "LEFTREF").First().StartBaseLine.X;
        double right = Word(letters, "RIGHTREF").Last().EndBaseLine.X;

        AssertSpan(Line(letters, "JRTL"), left, right);
        List<Letter> rtlLast = Line(letters, "JRTLEND");
        Assert.InRange(rtlLast.Last().EndBaseLine.X, right - .05D, right + .05D);
        Assert.True(rtlLast.First().StartBaseLine.X > left + 20D, "The last line of justified RTL text is not stretched.");

        List<Letter> ltrLast = Line(letters, "JLTREND");
        Assert.InRange(ltrLast.First().StartBaseLine.X, left - .05D, left + .05D);
        Assert.True(ltrLast.Last().EndBaseLine.X < right - 20D, "The last line of justified LTR text is not stretched.");

        AssertSpan(Line(letters, "DRTLEND"), left, right);
    }

    // The document default direction applies when the paragraph's style chain is empty, as with a missing style.
    [Fact]
    public void SaveAsPdf_DocumentDefaultDirectionAppliesWithoutAParagraphStyle() {
        var paragraphs = new List<W.Paragraph> {
            CreateParagraph("LEFTREF", false, W.JustificationValues.Left),
            CreateParagraph("RIGHTREF", false, W.JustificationValues.Right)
        };
        // Direct w:bidi="0" keeps the references left-to-right under the right-to-left default.
        foreach (W.Paragraph reference in paragraphs) reference.ParagraphProperties!.BiDi = new W.BiDi { Val = false };
        W.Paragraph missing = CreateParagraph("DEFAULTDIR", false, null);
        missing.ParagraphProperties!.ParagraphStyleId = new W.ParagraphStyleId { Val = "NoSuchStyle" };
        paragraphs.Add(missing);

        using var pdf = PdfPigDocument.Open(Convert("body", paragraphs, document => {
            W.DocDefaults defaults = document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.DocDefaults!;
            defaults.ParagraphPropertiesDefault ??= new W.ParagraphPropertiesDefault();
            defaults.ParagraphPropertiesDefault.ParagraphPropertiesBaseStyle ??= new W.ParagraphPropertiesBaseStyle();
            defaults.ParagraphPropertiesDefault.ParagraphPropertiesBaseStyle.BiDi = new W.BiDi();
        }));
        var letters = Assert.Single(pdf.GetPages()).Letters.ToList();
        double right = Word(letters, "RIGHTREF").Last().EndBaseLine.X;
        Assert.InRange(Word(letters, "DEFAULTDIR").Last().EndBaseLine.X, right - .05D, right + .05D);
    }

    // An inline shape group is placed without justification; in a justified right-to-left paragraph it keeps to the right.
    [Fact]
    public void SaveAsPdf_InlineShapeGroupInJustifiedRightToLeftParagraphKeepsToTheRight() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("RIGHTREF").ParagraphAlignment = WordParagraphAlignment.Right;
        WordParagraph paragraph = document.AddParagraph();
        paragraph._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        paragraph._paragraph.ParagraphProperties.BiDi = new W.BiDi();
        paragraph._paragraph.ParagraphProperties.Justification = new W.Justification { Val = W.JustificationValues.Both };
        paragraph.AddShapeGroup(new[] {
            new WordShapeGroupItem(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeLength.Points(0),
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(10)) { FillColor = OfficeIMO.OfficeColor.Parse("FF0000") },
            new WordShapeGroupItem(WordShapeType.Rectangle, OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(0),
                OfficeIMO.OfficeLength.Points(20), OfficeIMO.OfficeLength.Points(10)) { FillColor = OfficeIMO.OfficeColor.Parse("0000FF") }
        });

        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(400), OfficeIMO.OfficeLength.Points(700)),
            Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)),
            FontFamily = "Helvetica"
        }).ToBytes(new PdfSaveOptions { AllowLoss = true }));
        var page = Assert.Single(pdf.GetPages());
        double right = Word(page.Letters.ToList(), "RIGHTREF").Last().EndBaseLine.X;
        double groupRight = page.Paths.Where(path => path.IsFilled).Max(path => path.GetBoundingRectangle()!.Value.Right);
        Assert.InRange(groupRight, right - .5D, right + .5D);
    }

    // Text set around a paragraph's text box is justified like the rest of the paragraph.
    [Fact]
    public void SaveAsPdf_JustifiedParagraphWithATextBoxStillJustifiesItsText() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("LEFTREF").ParagraphAlignment = WordParagraphAlignment.Left;
        document.AddParagraph("RIGHTREF").ParagraphAlignment = WordParagraphAlignment.Right;
        WordParagraph paragraph = document.AddParagraph("JBOX " + Filler + " JBOXEND");
        paragraph.ParagraphAlignment = WordParagraphAlignment.Both;
        paragraph.AddTextBox("INNER", WordImageTextWrapping.InFrontOfText);

        var letters = ConvertDocument(document);
        double left = Word(letters, "LEFTREF").First().StartBaseLine.X;
        double right = Word(letters, "RIGHTREF").Last().EndBaseLine.X;
        AssertSpan(Line(letters, "JBOX"), left, right);
    }

    // List items that are justified or right-to-left keep their alignment.
    [Fact]
    public void SaveAsPdf_ListItemsKeepJustificationAndDirection() {
        using WordDocument document = WordDocument.Create();
        document.AddParagraph("LEFTREF").ParagraphAlignment = WordParagraphAlignment.Left;
        document.AddParagraph("RIGHTREF").ParagraphAlignment = WordParagraphAlignment.Right;
        WordList list = document.AddList(WordListStyle.Bulleted);
        WordParagraph justified = list.AddItem("JLIST " + Filler + " JLISTEND");
        justified.ParagraphAlignment = WordParagraphAlignment.Both;
        WordParagraph rightToLeft = list.AddItem("RTLITEM");
        rightToLeft._paragraph.ParagraphProperties!.BiDi = new W.BiDi();

        var letters = ConvertDocument(document);
        double right = Word(letters, "RIGHTREF").Last().EndBaseLine.X;
        Assert.InRange(Line(letters, "JLIST").Last().EndBaseLine.X, right - .05D, right + .05D);
        Assert.InRange(Word(letters, "RTLITEM").Last().EndBaseLine.X, right - .05D, right + .05D);
    }

    private static List<Letter> ConvertDocument(WordDocument document) {
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(400), OfficeIMO.OfficeLength.Points(700)),
            Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)),
            FontFamily = "Helvetica"
        }).ToBytes(new PdfSaveOptions { AllowLoss = true }));
        return Assert.Single(pdf.GetPages()).Letters.ToList();
    }

    private static void AssertSpan(List<Letter> line, double left, double right) {
        Assert.InRange(line.First().StartBaseLine.X, left - .05D, left + .05D);
        Assert.InRange(line.Last().EndBaseLine.X, right - .05D, right + .05D);
    }

    private static W.Paragraph CreateParagraph(string text, bool biDi, W.JustificationValues? alignment) {
        var properties = new W.ParagraphProperties();
        if (biDi) properties.BiDi = new W.BiDi();
        if (alignment.HasValue) properties.Justification = new W.Justification { Val = alignment.Value };
        properties.SpacingBetweenLines = new W.SpacingBetweenLines { After = "0" };
        return new W.Paragraph(properties, new W.Run(new W.Text(text) { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }));
    }

    private const string RightToLeftStyleId = "RightToLeftText";

    private static byte[] Convert(string container, IReadOnlyList<W.Paragraph> paragraphs, Action<WordDocument>? prepare = null) {
        using WordDocument document = WordDocument.Create();
        prepare?.Invoke(document);
        var main = document.OpenXmlDocument.MainDocumentPart!;
        main.StyleDefinitionsPart!.Styles!.Append(new W.Style(new W.StyleName { Val = RightToLeftStyleId },
            new W.StyleParagraphProperties(new W.BiDi())) { Type = W.StyleValues.Paragraph, StyleId = RightToLeftStyleId });
        switch (container) {
            case "body":
                document.AddParagraph("BODY");
                InsertAfter(main.Document.Body!.Elements<W.Paragraph>().Last(), paragraphs);
                break;
            case "cell":
                document.AddTable(1, 1);
                W.TableCell cell = main.Document.Body!.Descendants<W.TableCell>().Single();
                cell.RemoveAllChildren<W.Paragraph>();
                cell.Append(paragraphs);
                break;
            case "nestedcell":
                // A nested table sends the cell's paragraphs through the contained-flow planner.
                document.AddTable(1, 1);
                W.TableCell outer = main.Document.Body!.Descendants<W.TableCell>().Single();
                W.Table nested = (W.Table)main.Document.Body!.Elements<W.Table>().Single().CloneNode(true);
                outer.RemoveAllChildren<W.Paragraph>();
                outer.Append(paragraphs);
                outer.Append(nested, new W.Paragraph());
                break;
            case "header":
                document.AddParagraph("BODY"); _ = document.PageSetup.Header;
                WordParagraph first = document.PageSetup.GetHeader(WordHeaderFooterType.Default).AddParagraph("HEADER");
                InsertAfter(first._paragraph, paragraphs);
                break;
            case "textbox":
                WordTextBox box = document.AddTextBox("BOX", WordImageTextWrapping.InFrontOfText);
                box.Width = OfficeIMO.OfficeLength.Points(300);
                box.Height = OfficeIMO.OfficeLength.Points(400);
                InsertAfter(box.Paragraphs.Single()._paragraph, paragraphs);
                break;
        }
        return document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new PdfPageSize(OfficeIMO.OfficeLength.Points(400), OfficeIMO.OfficeLength.Points(700)),
            Margins = PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(30)),
            FontFamily = "Helvetica"
        }).ToBytes();
    }

    private static void InsertAfter(W.Paragraph anchor, IReadOnlyList<W.Paragraph> paragraphs) {
        W.Paragraph previous = anchor;
        foreach (W.Paragraph paragraph in paragraphs) previous = previous.InsertAfterSelf(paragraph);
    }

    private static List<Letter> Word(List<Letter> letters, string word) {
        for (int index = 0; index + word.Length <= letters.Count; index++) {
            if (!Enumerable.Range(0, word.Length).All(offset => letters[index + offset].Value == word[offset].ToString())) continue;
            bool startsWord = index == 0 || letters[index - 1].Value == " " || Math.Abs(letters[index - 1].StartBaseLine.Y - letters[index].StartBaseLine.Y) > .5D ||
                letters[index].StartBaseLine.X - letters[index - 1].EndBaseLine.X > .5D;
            int after = index + word.Length;
            bool endsWord = after == letters.Count || letters[after].Value == " " || Math.Abs(letters[after].StartBaseLine.Y - letters[index].StartBaseLine.Y) > .5D ||
                letters[after].StartBaseLine.X - letters[after - 1].EndBaseLine.X > .5D;
            if (startsWord && endsWord) return letters.GetRange(index, word.Length);
        }
        throw new Xunit.Sdk.XunitException($"Word {word} was not found.");
    }

    private static List<Letter> Line(List<Letter> letters, string word) {
        double baseline = Word(letters, word).First().StartBaseLine.Y;
        return letters.Where(letter => letter.Value != " " && Math.Abs(letter.StartBaseLine.Y - baseline) < .5D)
            .OrderBy(letter => letter.StartBaseLine.X).ToList();
    }
}
