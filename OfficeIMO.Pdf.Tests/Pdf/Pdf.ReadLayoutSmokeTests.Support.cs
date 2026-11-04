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
                PageSetup = {
                    Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(420), global::OfficeIMO.OfficeLength.Points(320)),
                    Margins = { Left = global::OfficeIMO.OfficeLength.Points(36), Right = global::OfficeIMO.OfficeLength.Points(36), Top = global::OfficeIMO.OfficeLength.Points(36), Bottom = global::OfficeIMO.OfficeLength.Points(36) }
                },
                DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10)
            })
            .Paragraph(p => p.AddText("First page table."))
            .Table(new[] {
                new[] { "Code", "Name", "Qty" },
                new[] { "A-100", "Alpha", "2" },
                new[] { "A-200", "Atlas", "4" }
            }, style: TableStyle())
            .PageBreak()
            .Paragraph(p => p.AddText("Second page marker."))
            .PageBreak()
            .Paragraph(p => p.AddText("Third page table."))
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
