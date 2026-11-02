using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates two small-format event badges as separate PDF pages suitable for individual printing.</summary>
    internal static class EventBadges {
        internal static void Create(string folder) {
            OfficeTheme theme = OfficeTheme.Default;
            OfficeColor accent = theme.Palette.Accent1;
            OfficeColor muted = theme.Palette.TextMuted;
            // Each badge is trimmed from a slightly larger sheet: 3 pt of bleed and crop marks in an 18 pt slug.
            const double width = 340, height = 255, slug = 18, bleed = 3, markLength = 12, markGap = 3;
            double sheetWidth = width + slug * 2, sheetHeight = height + slug * 2;
            var cropMarks = new List<PdfPageBackgroundShape>();
            foreach (double trimX in new[] { slug, slug + width }) {
                foreach (double trimY in new[] { slug, slug + height }) {
                    double outwardX = trimX == slug ? -1 : 1, outwardY = trimY == slug ? -1 : 1;
                    cropMarks.Add(CropMark(OfficeShape.Line(0, 0, markLength, 0),
                        outwardX < 0 ? trimX - markGap - markLength : trimX + markGap, trimY));
                    cropMarks.Add(CropMark(OfficeShape.Line(0, 0, 0, markLength),
                        trimX, outwardY < 0 ? trimY - markGap - markLength : trimY + markGap));
                }
            }
            PdfDocument.Create(pdf => pdf.Content(content => {
                Badge("Maya Ellis", "Facilitator", "Service design");
                content.PageBreak();
                Badge("Leon Brooks", "Participant", "Operations");
                void Badge(string name, string role, string team) {
                    content.Paragraph(p => p.FontSize(OfficeLength.Points(8.5)).Text("Northwind · Community workshop"), defaultColor: muted)
                        .HR(OfficeLength.Points(2), accent, OfficeLength.Points(8), OfficeLength.Points(14))
                        .Title(name)
                        .Paragraph(p => p.FontSize(OfficeLength.Points(13)).Bold(role, accent))
                        .Paragraph(p => p.Text(team), defaultColor: muted)
                        .PanelParagraph(p => p.Text("14 October 2026 · Copenhagen"));
                }
            }), new PdfCreateOptions {
                PageWidth = OfficeLength.Points(sheetWidth), PageHeight = OfficeLength.Points(sheetHeight),
                MarginLeft = OfficeLength.Points(slug + 24), MarginRight = OfficeLength.Points(slug + 24), MarginTop = OfficeLength.Points(slug + 24), MarginBottom = OfficeLength.Points(slug + 24), DefaultFontSize = OfficeLength.Points(10),
                PrintProductionPageBoxes = new PdfPrintProductionPageBoxes(PdfPageMargins.Uniform(OfficeLength.Points(slug)), PdfPageMargins.Uniform(OfficeLength.Points(slug - bleed))),
                PageBackgroundShapes = cropMarks
            })
                .Meta(title: "Workshop event badges", author: "OfficeIMO").Save(Path.Combine(folder, "example.pdf"));
        }

        private static PdfPageBackgroundShape CropMark(OfficeShape line, double x, double y) {
            line.StrokeColor = OfficeColor.Black;
            line.StrokeWidth = 0.25;
            return new PdfPageBackgroundShape(line, OfficeLength.Points(x), OfficeLength.Points(y));
        }
    }
}
