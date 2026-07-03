using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class PdfReadLayoutSmokeTests {
    private static byte[] BuildThreePageTablePdf() {
        return PdfDocument.Create(new PdfOptions {
                PageWidth = 420,
                PageHeight = 320,
                MarginLeft = 36,
                MarginRight = 36,
                MarginTop = 36,
                MarginBottom = 36,
                DefaultFontSize = 10
            })
            .Paragraph(p => p.Text("First page table."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "A-200", "Atlas", "4" }
            }, style: TableStyle())
            .PageBreak()
            .Paragraph(p => p.Text("Second page marker."))
            .PageBreak()
            .Paragraph(p => p.Text("Third page table."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "C-300", "Gamma", "5" },
                new[] { "C-400", "Comet", "7" }
            }, style: TableStyle())
            .ToBytes();
    }

    private static PdfTableStyle TableStyle() {
        return new PdfTableStyle {
            ColumnWidthPoints = new System.Collections.Generic.List<double?> { 70, 170, 60 },
            HeaderRowCount = 1,
            CellPaddingX = 6,
            CellPaddingY = 4
        };
    }

    private static string Normalize(string text) {
        return text.Replace(" ", string.Empty);
    }

    private static bool RowContains(string[] row, params string[] expectedTokens) {
        string rowText = NormalizeCsvText(string.Join(",", row));
        return expectedTokens.All(token => rowText.Contains(token, StringComparison.Ordinal));
    }

    private static string NormalizeCsvText(string text) {
        return new string(text.Where(ch => !char.IsWhiteSpace(ch) && ch != '"').ToArray());
    }
}
