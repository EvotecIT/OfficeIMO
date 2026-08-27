using System.IO;
using System.Linq;
using OfficeIMO.Excel;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.OpenDocument;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenDocument;
using Xunit;

namespace OfficeIMO.OpenDocument.Converters.Tests;

public sealed class OpenDocumentTextFormattingConversionTests {
    [Fact]
    public void WordAndOdtPreserveNativeDecorationScriptAndCapsSemantics() {
        using WordDocument source = WordDocument.Create();
        WordParagraph run = source.AddParagraph("Styled");
        run.Underline = WordUnderlineStyle.WavyDouble;
        run.DoubleStrike = true;
        run.VerticalTextAlignment = WordVerticalTextPosition.Superscript;
        run.CapsStyle = WordCapsStyle.SmallCaps;

        OdtDocument persisted = OdtDocument.Load(new MemoryStream(source.ToOpenDocument().ToBytes()));
        OdtSpan odtRun = Assert.Single(Assert.Single(persisted.Paragraphs).Spans);
        Assert.Equal(OdfTextDecorationStyle.Wave, odtRun.UnderlineStyle);
        Assert.Equal(OdfTextDecorationType.Double, odtRun.UnderlineType);
        Assert.Equal(OdfTextDecorationType.Double, odtRun.LineThroughType);
        Assert.Equal(OdfTextPosition.Superscript, odtRun.TextPosition);
        Assert.True(odtRun.SmallCaps);

        using WordDocument roundTrip = persisted.ToWordDocument();
        WordRunSnapshot converted = Assert.Single(Assert.Single(roundTrip.CreateInspectionSnapshot().Sections
            .SelectMany(section => section.Elements).OfType<WordParagraphSnapshot>()).Runs);
        Assert.Equal(WordUnderlineStyle.WavyDouble, converted.UnderlineStyle);
        Assert.True(converted.DoubleStrike);
        Assert.Equal(nameof(WordVerticalTextPosition.Superscript), converted.VerticalTextAlignment);
        Assert.Equal(nameof(WordCapsStyle.SmallCaps), converted.CapsStyle);
    }

    [Fact]
    public void PowerPointAndOdpPreserveNativeDecorationScriptAndCapsSemantics() {
        using PowerPointPresentation source = PowerPointPresentation.Create(new MemoryStream(), new PowerPointCreateOptions());
        PowerPointTextRun run = source.AddSlide().AddTextBoxPoints("Styled", 10, 10, 200, 40)
            .Paragraphs[0].Runs[0];
        run.UnderlineStyle = PowerPointUnderlineStyle.WavyDouble;
        run.StrikeStyle = PowerPointStrikeStyle.Double;
        run.SetSubscript();
        run.Capitalization = PowerPointCapitalization.SmallCaps;

        OdpPresentation persisted = OdpPresentation.Load(new MemoryStream(source.ToOpenDocument().ToBytes()));
        OdpRun odpRun = Assert.Single(Assert.IsType<OdpTextBox>(Assert.Single(persisted.Slides[0].Shapes))
            .Paragraphs[0].Runs);
        Assert.Equal(OdfTextDecorationStyle.Wave, odpRun.UnderlineStyle);
        Assert.Equal(OdfTextDecorationType.Double, odpRun.UnderlineType);
        Assert.Equal(OdfTextDecorationType.Double, odpRun.LineThroughType);
        Assert.Equal(OdfTextPosition.Subscript, odpRun.TextPosition);
        Assert.True(odpRun.SmallCaps);

        using PowerPointPresentation roundTrip = persisted.ToPowerPointPresentation();
        PowerPointTextRun converted = roundTrip.Slides[0].TextBoxes.Single().Paragraphs[0].Runs[0];
        Assert.Equal(PowerPointUnderlineStyle.WavyDouble, converted.UnderlineStyle);
        Assert.Equal(PowerPointStrikeStyle.Double, converted.StrikeStyle);
        Assert.True(converted.BaselinePercent < 0D);
        Assert.Equal(PowerPointCapitalization.SmallCaps, converted.Capitalization);
    }

    [Fact]
    public void ExcelAndOdsPreserveCellDecorationAndScriptAndApplyDisplayCase() {
        using ExcelDocument source = ExcelDocument.Create();
        ExcelCell cell = source.AddWorksheet("Data").CellAt(1, 1)
            .SetValue("Styled")
            .SetUnderline(ExcelUnderlineStyle.Double)
            .SetStrikethrough()
            .SetSuperscript();

        OdsDocument persisted = OdsDocument.Load(new MemoryStream(source.ToOpenDocument().ToBytes()));
        OdsCell odsCell = persisted.Sheets.Single().Cell(0, 0);
        Assert.Equal(OdfTextDecorationStyle.Solid, odsCell.UnderlineStyle);
        Assert.Equal(OdfTextDecorationType.Double, odsCell.UnderlineType);
        Assert.True(odsCell.StrikeThrough);
        Assert.Equal(OdfTextPosition.Superscript, odsCell.TextPosition);

        odsCell.SetString("MiXeD");
        odsCell.TextTransform = OdfTextTransform.Lowercase;
        using ExcelDocument roundTrip = OdsDocument.Load(new MemoryStream(persisted.ToBytes())).ToExcelDocument();
        ExcelCellSnapshot converted = Assert.Single(roundTrip.CreateInspectionSnapshot().Worksheets.Single().Cells);
        Assert.Equal("mixed", converted.Value);
        Assert.Equal(ExcelUnderlineStyle.Double, converted.Style!.UnderlineStyle);
        Assert.True(converted.Style.Strikethrough);
        Assert.Equal(ExcelVerticalTextAlignment.Superscript, converted.Style.VerticalTextAlignment);
    }
}
