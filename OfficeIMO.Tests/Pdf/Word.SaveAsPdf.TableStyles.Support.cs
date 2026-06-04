using DocumentFormat.OpenXml.Wordprocessing;
using System;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Word;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests;

public partial class Word {
    private static void ConfigureMarginTable(WordTable table, string label) {
        WordTableCell cell = table.Rows[0].Cells[0];
        cell.Width = 2880;
        cell.WidthType = TableWidthUnitValues.Dxa;
        cell.Paragraphs[0].Text = label;
    }

    private static void ConfigurePlacementTable(WordTable table, string label, TableRowAlignmentValues alignment) {
        table.Alignment = alignment;
        foreach (WordTableCell cell in table.Rows[0].Cells) {
            cell.Width = 1440;
            cell.WidthType = TableWidthUnitValues.Dxa;
        }

        table.Rows[0].Cells[0].Paragraphs[0].Text = label;
        table.Rows[0].Cells[1].Paragraphs[0].Text = "Value";
    }

    private static void ConfigureCellSpacingTable(WordTable table, string left, string right) {
        foreach (WordTableCell cell in table.Rows[0].Cells) {
            cell.Width = 1440;
            cell.WidthType = TableWidthUnitValues.Dxa;
        }

        table.Rows[0].Cells[0].Paragraphs[0].Text = left;
        table.Rows[0].Cells[1].Paragraphs[0].Text = right;
    }

    private static PdfCore.PdfTableStyle CreateNativeTableStyleForTest(WordTable table) {
        MethodInfo method = typeof(WordPdfConverterExtensions).GetMethod("CreateNativeTableStyle", BindingFlags.NonPublic | BindingFlags.Static)!;
        return Assert.IsType<PdfCore.PdfTableStyle>(method.Invoke(null, new object?[] { table, table.Rows.Count, null }));
    }
}
