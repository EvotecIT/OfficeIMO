using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordPaperTablesTests {
    [Fact]
    public void AddCountTable_ContentSized_AlignsNumericHeaderAndValues() {
        using WordDocument document = WordDocument.Create();
        WordTable table = new WordPaperStyle().AddCountTable(document, new[] {
            new KeyValuePair<string, string>("Generated recommendations", "12"),
            new KeyValuePair<string, string>("Approved", "2")
        });
        Assert.Equal(WordTableWidthKind.Auto, table.Width?.Kind);
        Assert.All(table.Rows, row => {
            Assert.All(row.Cells[0].Paragraphs, paragraph => Assert.Equal(WordParagraphAlignment.Left, paragraph.ParagraphAlignment));
            Assert.All(row.Cells[1].Paragraphs, paragraph => Assert.Equal(WordParagraphAlignment.Right, paragraph.ParagraphAlignment));
        });
    }

    [Fact]
    public void AddKeyValueTable_NumbersAndText_KeepValuesLeftAndCoverHeaderAbsent() {
        using WordDocument document = WordDocument.Create();
        WordTable table = new WordPaperStyle().AddKeyValueTable(document, new[] {
            new KeyValuePair<string, string>("Hosts", "1"),
            new KeyValuePair<string, string>("Compliance", "100%"),
            new KeyValuePair<string, string>("Coverage", "62.5%"),
            new KeyValuePair<string, string>("Scope", "corp.example")
        }, hasHeader: false);
        Assert.Equal(4, table.Rows.Count);
        Assert.False(table.RepeatHeaderRowAtTheTopOfEachPage);
        Assert.All(table.Rows.SelectMany(row => row.Cells).SelectMany(cell => cell.Paragraphs),
            paragraph => Assert.Equal(WordParagraphAlignment.Left, paragraph.ParagraphAlignment));
    }

    [Fact]
    public void ApplyTable_NumericColumn_AlignsHeaderWithValues() {
        using WordDocument document = WordDocument.Create();
        WordTable table = document.AddTable(2, 3);
        table.Rows[0].Cells[0].Paragraphs[0].Text = "Area";
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Count";
        table.Rows[0].Cells[2].Paragraphs[0].Text = "State";
        table.Rows[1].Cells[0].Paragraphs[0].Text = "Identity";
        table.Rows[1].Cells[1].Paragraphs[0].Text = "12";
        table.Rows[1].Cells[2].Paragraphs[0].Text = "Review";
        new WordPaperStyle().ApplyTable(table);
        Assert.All(table.Rows, row => Assert.Equal(WordParagraphAlignment.Right, row.Cells[1].Paragraphs[0].ParagraphAlignment));
        Assert.All(table.Rows, row => Assert.Equal(WordParagraphAlignment.Left, row.Cells[2].Paragraphs[0].ParagraphAlignment));
    }

    [Fact]
    public void AddPageChrome_PageFields_UseSameMutedTypographyAsFooterText() {
        using WordDocument document = WordDocument.Create();
        var paper = new WordPaperStyle();
        paper.AddPageChrome(document, "Report", "Product", "Scope");
        using MemoryStream stream = new();
        document.Save(stream);
        stream.Position = 0;
        using WordprocessingDocument package = WordprocessingDocument.Open(stream, false);
        Assert.NotEmpty(package.MainDocumentPart!.FooterParts);
        foreach (FooterPart part in package.MainDocumentPart.FooterParts) {
            var runs = part.Footer!.Descendants<Run>().ToArray();
            Assert.Contains(part.Footer.Descendants<SimpleField>(), field => field.Instruction?.Value?.Contains("PAGE") == true);
            Assert.All(runs, run => {
                Assert.Equal("16", run.RunProperties?.FontSize?.Val?.Value);
                Assert.Equal(paper.Muted, run.RunProperties?.Color?.Val?.Value);
            });
        }
    }
}
