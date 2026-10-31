using System;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using OfficeIMO.Tests.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class PdfReadLayoutSmokeTests {
    private static byte[] BuildThreePageTablePdf() {
        return PdfDocument.Create(new PdfCreateOptions {
                PageWidth = global::OfficeIMO.OfficeLength.Points(420),
                PageHeight = global::OfficeIMO.OfficeLength.Points(320),
                MarginLeft = global::OfficeIMO.OfficeLength.Points(36),
                MarginRight = global::OfficeIMO.OfficeLength.Points(36),
                MarginTop = global::OfficeIMO.OfficeLength.Points(36),
                MarginBottom = global::OfficeIMO.OfficeLength.Points(36),
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
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
            ColumnWidths = new System.Collections.Generic.List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(70), global::OfficeIMO.OfficeLength.Points(170), global::OfficeIMO.OfficeLength.Points(60) },
            HeaderRowCount = 1,
            CellPaddingX = global::OfficeIMO.OfficeLength.Points(6),
            CellPaddingY = global::OfficeIMO.OfficeLength.Points(4)
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
