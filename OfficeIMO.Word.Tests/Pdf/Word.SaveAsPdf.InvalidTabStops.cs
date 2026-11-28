using DocumentFormat.OpenXml;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using W = DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>
/// A tab stop whose position isn't a number has no place on the line, and an alignment or leader token the schema
/// doesn't define reads as missing: PDF export lays the paragraph out as if those values weren't there.
/// </summary>
public sealed class WordPdfInvalidTabStopTests {
    private const string WordNamespace = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    [Fact]
    public void UnreadableTabStopValuesReadAsMissing() {
        double expected = SecondWordLeft(invalid: false);

        double actual = SecondWordLeft(invalid: true);

        Assert.Equal(expected, actual, 2);
    }

    private static double SecondWordLeft(bool invalid) {
        using var document = WordDocument.Create();
        // A style stop that the paragraph clears: the clear still masks it.
        document.OpenXmlDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(new W.Style(
            new W.StyleName { Val = "Tabbed" },
            new W.StyleParagraphProperties(new W.Tabs(new W.TabStop { Val = W.TabStopValues.Left, Position = 1440 }))) {
            Type = W.StyleValues.Paragraph, StyleId = "Tabbed"
        });
        WordParagraph paragraph = document.AddParagraph("FIRST");
        paragraph._paragraph.Append(new W.Run(new W.TabChar(), new W.Text("SECOND")));
        var stops = new List<W.TabStop> {
            new() { Val = W.TabStopValues.Clear, Position = 1440 },
            new() { Val = W.TabStopValues.Right, Position = 6000 }
        };
        if (invalid) {
            stops[1].SetAttribute(new OpenXmlAttribute("w", "leader", WordNamespace, "bogus"));
            var unplaced = new W.TabStop { Val = W.TabStopValues.Left };
            unplaced.SetAttribute(new OpenXmlAttribute("w", "pos", WordNamespace, "bad"));
            stops.Insert(0, unplaced);
            var unaligned = new W.TabStop { Position = 8000 };
            unaligned.SetAttribute(new OpenXmlAttribute("w", "val", WordNamespace, "sideways"));
            stops.Add(unaligned);
        } else {
            stops.Add(new W.TabStop { Position = 8000 });
        }
        paragraph._paragraph.ParagraphProperties = new W.ParagraphProperties(new W.ParagraphStyleId { Val = "Tabbed" }, new W.Tabs(stops));

        PdfDocumentConversionResult result = document.ConvertToPdfWithEvidence(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic()
        });
        using var pdf = PdfPigDocument.Open(result.ToBytes());
        return Assert.Single(Assert.Single(pdf.GetPages()).GetWords(), word => word.Text == "SECOND").BoundingBox.Left;
    }
}
