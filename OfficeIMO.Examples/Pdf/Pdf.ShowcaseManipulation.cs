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

            PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Source A", PdfAlign.Left)
                .Paragraph(p => p.Text("This page is created first, then merged with Source B through the fluent PdfDocument API."))
                .Table(new[] {
                    new[] { "Capability", "Status", "Notes" },
                    new[] { "Create", "Ready", "Dependency-free writer surface" },
                    new[] { "Inspect", "Ready", "Metadata, pages, links, outlines, blockers" },
                    new[] { "Extract", "Ready", "Page ranges and reordered pages" }
                }, style: UtilityTableStyle())
                .PageBreak()
                .H1("Source A \u2013 details", PdfAlign.Left)
                .PanelParagraph(
                    p => p.Bold("Page two is extracted later. ").Text("This proves the example is not just visual; it also exercises read/rewrite operations."),
                    new PdfPanelStyle {
                        Background = PdfColor.FromRgb(248, 250, 252),
                        BorderColor = PdfColor.FromRgb(148, 163, 184),
                        BorderWidth = 0.7,
                        PaddingX = 10,
                        PaddingY = 8
                    })), StandardOptions("OfficeIMO.Pdf manipulation source A"))
                .Meta(title: "OfficeIMO.Pdf Manipulation Source A", author: "OfficeIMO")
                .Save(sourceAPath);

            PdfDocument.Create(pdf => pdf.Content(content => content
                .H1("Source B", PdfAlign.Left)
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
                }, style: UtilityTableStyle())), StandardOptions("OfficeIMO.Pdf manipulation source B"))
                .Meta(title: "OfficeIMO.Pdf Manipulation Source B", author: "OfficeIMO")
                .Save(sourceBPath);

            PdfDocument merged = PdfDocument.Load(sourceAPath)
                .MergeWith(sourceBPath);
            merged.Save(mergedPath);
            merged.Pages.Extract(PdfPageRange.From(2, 2))
                .Save(extractedPath);
            // The sources carry no page numbers; the merged pack is numbered once, after the merge,
            // on the left margin and baseline the source footers would have used.
            PdfDocument stamped = merged;
            int pageCount = merged.Pages.Count;
            for (int page = 1; page <= pageCount; page++) {
                stamped = stamped.Stamp.Text(
                    "OfficeIMO.Pdf examples \u00B7 Page " + page + " of " + pageCount,
                    new PdfTextStampOptions {
                        PageNumbers = new[] { page },
                        X = 72,
                        Y = 36,
                        Font = PdfStandardFont.Helvetica,
                        FontSize = 8.5,
                        Color = PdfColor.FromOfficeColor(OfficeTheme.Default.Palette.TextMuted)
                    });
            }
            stamped.Save(stampedPath);
            File.WriteAllText(extractedTextPath, stamped.Read().Text);

            if (open) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo { FileName = stampedPath, UseShellExecute = true });
            }
        }

        private static PdfCreateOptions StandardOptions(string header) {
            return new PdfCreateOptions {
                HeaderFormat = header,
                HeaderAlign = PdfAlign.Left,
                ShowHeader = true,
                CreateOutlineFromHeadings = true
            };
        }

        private static PdfTableStyle UtilityTableStyle() {
            var style = PdfTheme.FromOfficeTheme(OfficeTheme.Default).TableStyle!;
            style.CellPaddingX = 7;
            style.CellPaddingY = 5;
            style.CellPaddingTop = 5;
            style.CellPaddingBottom = 5;
            style.FontSize = 9.5;
            style.LineHeight = 1.18;
            style.AutoFitColumns = true;
            return style;
        }
    }
}
