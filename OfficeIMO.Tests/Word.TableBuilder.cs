using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void Test_TableBuilder_BuildsWordTable() {
        List<List<Action<TableCell>>> structure = new() {
            new() {
                cell => cell.Append(new Paragraph(new Run(new Text("A")) )),
                cell => cell.Append(new Paragraph(new Run(new Text("B")) ))
            },
            new() {
                cell => cell.Append(new Paragraph(new Run(new Text("C")) )),
                cell => cell.Append(new Paragraph(new Run(new Text("D")) ))
            }
        };

        Table table = TableBuilder.Build(structure);
        var rows = table.Elements<TableRow>().ToList();
        Assert.Equal(2, rows.Count);
        var firstRowCells = rows[0].Elements<TableCell>().ToList();
        Assert.Equal("A", firstRowCells[0].InnerText);
        Assert.Equal("B", firstRowCells[1].InnerText);
    }

    [Fact]
    public void Test_TableBuilder_MapsWordTable() {
        string path = Path.Combine(_directoryWithFiles, "TableBuilder.docx");
        using (WordDocument document = WordDocument.Create(path)) {
            WordTable wt = document.AddTable(2, 2);
            wt.Rows[0].Cells[0].Paragraphs[0].Text = "A";
            wt.Rows[0].Cells[1].Paragraphs[0].Text = "B";
            wt.Rows[1].Cells[0].Paragraphs[0].Text = "C";
            wt.Rows[1].Cells[1].Paragraphs[0].Text = "D";
            document.Save();

            var mapped = TableBuilder.Map(wt).ToList();
            Assert.Equal(2, mapped.Count);
            Assert.Equal("A", mapped[0][0].Paragraphs[0].Text);
            Assert.Equal("D", mapped[1][1].Paragraphs[0].Text);
        }
    }
}
