using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates two small-format event badges as separate PDF pages suitable for individual printing.</summary>
    internal static class EventBadges {
        internal static void Create(string folder) {
            OfficeTheme theme = OfficeTheme.Default;
            PdfColor accent = PdfColor.FromOfficeColor(theme.Palette.Accent1);
            PdfColor muted = PdfColor.FromOfficeColor(theme.Palette.TextMuted);
            const double width = 340, height = 255;
            PdfDocument.Create(pdf => pdf.Content(content => {
                Badge("Maya Ellis", "Facilitator", "Service design");
                content.PageBreak();
                Badge("Leon Brooks", "Participant", "Operations");
                void Badge(string name, string role, string team) {
                    content.Paragraph(p => p.FontSize(8.5).Text("Northwind · Community workshop"), defaultColor: muted)
                        .HR(2, accent, 8, 14)
                        .Title(name)
                        .Paragraph(p => p.FontSize(13).Bold(role, accent))
                        .Paragraph(p => p.Text(team), defaultColor: muted)
                        .PanelParagraph(p => p.Text("14 October 2026 · Copenhagen"));
                }
            }), new PdfCreateOptions { PageWidth = width, PageHeight = height, MarginLeft = 24, MarginRight = 24, MarginTop = 24, MarginBottom = 24, DefaultFontSize = 10 })
                .Meta(title: "Workshop event badges", author: "OfficeIMO").Save(Path.Combine(folder, "example.pdf"));
        }
    }
}
