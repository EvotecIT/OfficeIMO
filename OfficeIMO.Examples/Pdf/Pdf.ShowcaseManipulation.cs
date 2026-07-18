using OfficeIMO.Pdf;
using System.IO;

namespace OfficeIMO.Examples.Pdf {
    internal static class ShowcaseManipulationPdf {
        public static void Example_Pdf_ShowcaseManipulation(string folderPath, bool open = false) {
            string sourceAPath = Path.Combine(folderPath, "Pdf.Showcase.Manipulation.SourceA.pdf");
            string sourceBPath = Path.Combine(folderPath, "Pdf.Showcase.Manipulation.SourceB.pdf");
            string mergedPath = Path.Combine(folderPath, "Pdf.Showcase.Manipulation.Merged.pdf");
            string extractedPath = Path.Combine(folderPath, "Pdf.Showcase.Manipulation.ExtractedPage2.pdf");
            string stampedPath = Path.Combine(folderPath, "Pdf.Showcase.Manipulation.Stamped.pdf");
            string extractedTextPath = Path.Combine(folderPath, "Pdf.Showcase.Manipulation.ExtractedText.txt");

            PdfDocument.Create(StandardOptions("OfficeIMO.Pdf manipulation source A"))
                .Meta(title: "OfficeIMO.Pdf Manipulation Source A", author: "OfficeIMO")
                .H1("Source A", PdfAlign.Left, PdfColor.FromRgb(15, 23, 42))
                .Paragraph(p => p.Text("This page is created first, then merged with Source B through the fluent PdfDocument API."))
                .Table(new[] {
                    new[] { "Capability", "Status", "Notes" },
                    new[] { "Create", "Ready", "Dependency-free writer surface" },
                    new[] { "Inspect", "Ready", "Metadata, pages, links, outlines, blockers" },
                    new[] { "Extract", "Ready", "Page ranges and reordered pages" }
                }, style: UtilityTableStyle())
                .PageBreak()
                .H1("Source A - details", PdfAlign.Left, PdfColor.FromRgb(15, 23, 42))
                .PanelParagraph(
                    p => p.Bold("Page two is extracted later. ").Text("This proves the example is not just visual; it also exercises read/rewrite operations."),
                    new PanelStyle {
                        Background = PdfColor.FromRgb(248, 250, 252),
                        BorderColor = PdfColor.FromRgb(148, 163, 184),
                        BorderWidth = 0.7,
                        PaddingX = 10,
                        PaddingY = 8
                    })
                .Save(sourceAPath);

            PdfDocument.Create(StandardOptions("OfficeIMO.Pdf manipulation source B"))
                .Meta(title: "OfficeIMO.Pdf Manipulation Source B", author: "OfficeIMO")
                .H1("Source B", PdfAlign.Left, PdfColor.FromRgb(15, 23, 42))
                .Paragraph(p => p.Text("A second input document is merged into the final pack, then a stamp is applied to the combined output."))
                .Bullets(new[] {
                    "Merge reads source pages and rewrites a new document.",
                    "Stamp adds a content stream to selected pages.",
                    "Text extraction writes a sidecar text file for PowerShell-friendly workflows."
                })
                .Table(new[] {
                    new[] { "Operation", "Output" },
                    new[] { "MergeWith", Path.GetFileName(mergedPath) },
                    new[] { "Pages.Extract", Path.GetFileName(extractedPath) },
                    new[] { "Stamp.Text", Path.GetFileName(stampedPath) },
                    new[] { "Read.Text", Path.GetFileName(extractedTextPath) }
                }, style: UtilityTableStyle())
                .Save(sourceBPath);

            PdfDocument merged = PdfDocument.Open(sourceAPath)
                .MergeWith(sourceBPath);
            merged.Save(mergedPath);
            merged.Pages.Extract(PdfPageRange.From(2, 2))
                .Save(extractedPath);
            PdfDocument stamped = merged.Stamp.Text(
                "OfficeIMO.Pdf",
                new PdfTextStampOptions {
                    X = 386,
                    Y = 30,
                    Font = PdfStandardFont.HelveticaBold,
                    FontSize = 11,
                    Color = PdfColor.FromRgb(14, 116, 144)
                });
            stamped.Save(stampedPath);
            File.WriteAllText(extractedTextPath, stamped.Read.Text());

            if (open) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = stampedPath, UseShellExecute = true });
            }
        }

        private static PdfOptions StandardOptions(string header) {
            return new PdfOptions {
                DefaultFont = PdfStandardFont.Helvetica,
                DefaultFontSize = 10,
                DefaultTextColor = PdfColor.FromRgb(31, 41, 55),
                HeaderFont = PdfStandardFont.Helvetica,
                HeaderFontSize = 8,
                HeaderFormat = header,
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                FooterFont = PdfStandardFont.Helvetica,
                FooterFontSize = 8,
                FooterFormat = "OfficeIMO.Pdf examples - page {page}/{pages}",
                FooterAlign = PdfAlign.Right,
                ShowPageNumbers = true,
                CreateOutlineFromHeadings = true
            };
        }

        private static PdfTableStyle UtilityTableStyle() {
            return new PdfTableStyle {
                HeaderFill = PdfColor.FromRgb(15, 23, 42),
                HeaderTextColor = PdfColor.White,
                RowStripeFill = PdfColor.FromRgb(248, 250, 252),
                BorderColor = PdfColor.FromRgb(203, 213, 225),
                BorderWidth = 0.5,
                CellPaddingX = 7,
                CellPaddingY = 5,
                FontSize = 9.5,
                LineHeight = 1.18,
                AutoFitColumns = true
            };
        }
    }
}
