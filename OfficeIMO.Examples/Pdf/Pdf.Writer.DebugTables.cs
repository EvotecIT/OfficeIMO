using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class WriterDebugTables {
        public static void Example_Pdf_TableDebug(string folderPath, bool open = false) {
            string path = Path.Combine(folderPath, "Pdf.TableDebug.pdf");

            var rows = new[] {
                new [] { "Item", "Qty", "Cost" },
                new [] { "Pencils", "3", "$1.20" },
                new [] { "Notebooks", "2", "$4.00" },
                new [] { "Folders", "5", "$2.50" },
                new [] { "Markers", "8", "$6.40" },
            };

            var debug = new PdfDebugOptions {
                ShowContentArea = true,
                ShowTableBaselines = true,
                ShowTableRowBoxes = true,
                ShowTableColumnGuides = true
            };

            var opts = new PdfOptions { Debug = debug };

            var styleBase = new PdfTableStyle {
                HeaderFill = PdfColor.LightGray,
                BorderColor = PdfColor.FromRgb(210,210,210),
                BorderWidth = 0.5,
                CellPaddingX = 6,
                CellPaddingY = 13.5
            };

            PdfDoc.Create(opts)
                .H1("Table Debug — Offset -1.0", PdfAlign.Left)
                .Table(rows, PdfAlign.Left, new PdfTableStyle {
                    HeaderFill = styleBase.HeaderFill,
                    BorderColor = styleBase.BorderColor,
                    BorderWidth = styleBase.BorderWidth,
                    CellPaddingX = styleBase.CellPaddingX,
                    CellPaddingY = styleBase.CellPaddingY,
                    RowBaselineOffset = -1.0
                })
                .Paragraph(p => p.Text(" "))
                .H1("Table Debug — Offset 0.0", PdfAlign.Left)
                .Table(rows, PdfAlign.Left, new PdfTableStyle {
                    HeaderFill = styleBase.HeaderFill,
                    BorderColor = styleBase.BorderColor,
                    BorderWidth = styleBase.BorderWidth,
                    CellPaddingX = styleBase.CellPaddingX,
                    CellPaddingY = styleBase.CellPaddingY,
                    RowBaselineOffset = 0.0
                })
                .Paragraph(p => p.Text(" "))
                .H1("Table Debug — Offset +1.0", PdfAlign.Left)
                .Table(rows, PdfAlign.Left, new PdfTableStyle {
                    HeaderFill = styleBase.HeaderFill,
                    BorderColor = styleBase.BorderColor,
                    BorderWidth = styleBase.BorderWidth,
                    CellPaddingX = styleBase.CellPaddingX,
                    CellPaddingY = styleBase.CellPaddingY,
                    RowBaselineOffset = 1.0
                })
                .Save(path);

            if (open) System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = path, UseShellExecute = true });
        }
    }
}

