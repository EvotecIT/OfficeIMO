using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Globalization;
using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Letter = UglyToad.PdfPig.Content.Letter;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

[Collection("WordTests")]
public sealed class WordNativeCompressionAdmissionTests {
    [Theory]
    [InlineData("paragraph")]
    [InlineData("heading")]
    [InlineData("table")]
    public void SaveAsPdf_SourceJustificationAdmitsAnOtherwiseWrappedWordThroughItsRealCarrier(string owner) {
        List<Letter> natural = Assert.Single(Render(owner, W.JustificationValues.Left, 240D));
        double width = natural[natural.Count - 1].EndBaseLine.X - natural[0].StartBaseLine.X - 0.5D;
        Assert.True(Render(owner, W.JustificationValues.Left, width).Count > 1);
        List<Letter> actual = Assert.Single(Render(owner, W.JustificationValues.Both, width));
        Assert.Equal("Alphabetagamma", string.Concat(actual.Select(letter => letter.Value)));
        Assert.True(actual[actual.Count - 1].EndBaseLine.X - actual[0].StartBaseLine.X <
            natural[natural.Count - 1].EndBaseLine.X - natural[0].StartBaseLine.X - 0.1D);
    }

    private static List<List<Letter>> Render(string owner, W.JustificationValues alignment, double width) {
        using WordDocument document = WordDocument.Create();
        WordParagraph target;
        if (owner == "table") {
            WordTable table = document.AddTable(1, 1);
            table.LayoutMode = WordTableLayoutMode.Fixed;
            string twips = Math.Round(width * 20D).ToString(CultureInfo.InvariantCulture);
            table._tableProperties!.TableWidth = new W.TableWidth { Type = W.TableWidthUnitValues.Dxa, Width = twips };
            WordTableCell cell = table.Rows[0].Cells[0];
            cell.Width = int.Parse(twips, CultureInfo.InvariantCulture);
            cell.WidthType = WordTableWidthUnit.Dxa;
            cell.MarginLeftWidth = 0;
            cell.MarginRightWidth = 0;
            target = cell.Paragraphs[0];
        } else {
            target = document.AddParagraph();
            if (owner == "heading") target.SetStyle(WordParagraphStyles.Heading1);
        }
        target._paragraph.RemoveAllChildren<W.Run>();
        target._paragraph.ParagraphProperties ??= new W.ParagraphProperties();
        var properties = target._paragraph.ParagraphProperties;
        properties.Justification = new W.Justification { Val = alignment };
        properties.SpacingBetweenLines = new W.SpacingBetweenLines { Before = "0", After = "0", Line = "280", LineRule = W.LineSpacingRuleValues.Exact };
        target._paragraph.Append(new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Helvetica", HighAnsi = "Helvetica" },
            new W.FontSize { Val = "24" }), new W.Text("Alpha beta gam") { Space = SpaceProcessingModeValues.Preserve }));
        target._paragraph.Append(new W.Run(new W.RunProperties(new W.RunFonts { Ascii = "Helvetica", HighAnsi = "Helvetica" },
            new W.Color { Val = "CC0000" }, new W.FontSize { Val = "24" }), new W.Text("ma")));
        Assert.Empty(document.ValidateDocument());
        string authored = document.OpenXmlDocument.MainDocumentPart!.Document.OuterXml;
        byte[] bytes = document.ToPdfDocumentResult(new WordToPdfOptions { IncludePageNumbers = false,
            PageSize = new PageSize(owner == "table" ? 300D : width + 60D, 300D), Margins = PageMargins.Uniform(30D), FontFamily = "Helvetica" }).ToBytes();
        Assert.Equal(authored, document.OpenXmlDocument.MainDocumentPart.Document.OuterXml);
        Assert.Empty(document.ValidateDocument());
        string? output = Environment.GetEnvironmentVariable("OFFICEIMO_PDF_ENGINE_PROOF_OUTPUT");
        if (!string.IsNullOrWhiteSpace(output)) {
            Directory.CreateDirectory(output!);
            File.WriteAllBytes(Path.Combine(output!, "word-" + owner + "-" + alignment + "-"
                + width.ToString("F3", CultureInfo.InvariantCulture) + "-" + Guid.NewGuid().ToString("N") + ".pdf"), bytes);
        }
        using var pdf = PdfPigDocument.Open(bytes);
        return pdf.GetPages().SelectMany(page => page.Letters.Where(letter => !string.IsNullOrWhiteSpace(letter.Value))
            .GroupBy(letter => Math.Round(letter.StartBaseLine.Y, 2)).Select(line => line.OrderBy(letter => letter.StartBaseLine.X).ToList())).ToList();
    }
}
