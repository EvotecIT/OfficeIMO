using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates two small-format event badges as separate PDF pages suitable for individual printing.</summary>
    internal static class EventBadges {
        internal static void Create(string folder) {
            var navy = global::OfficeIMO.OfficeColor.FromRgb(23, 54, 93);
            PdfDocument.Create(pdf => pdf.Content(content => {
                Badge("Maya Ellis", "FACILITATOR", "Service design");
                content.PageBreak();
                Badge("Leon Brooks", "PARTICIPANT", "Operations");
                void Badge(string name, string role, string team) {
                    content.Paragraph(p => p.Text("NORTHWIND / COMMUNITY WORKSHOP"))
                        .HR(global::OfficeIMO.OfficeLength.Points(2), navy, global::OfficeIMO.OfficeLength.Points(12), global::OfficeIMO.OfficeLength.Points(18))
                        .H1(name, PdfAlign.Left, navy)
                        .H2(role, PdfAlign.Left, navy)
                        .Paragraph(p => p.Text(team))
                        .PanelParagraph(p => p.Text("14 OCT 2026  |  COPENHAGEN"),
                            new PdfPanelStyle { Background = global::OfficeIMO.OfficeColor.FromRgb(234, 241, 251), PaddingX = 10, PaddingY = 10 });
                }
            }), new PdfCreateOptions { PageWidth = global::OfficeIMO.OfficeLength.Points(340), PageHeight = global::OfficeIMO.OfficeLength.Points(255), MarginLeft = global::OfficeIMO.OfficeLength.Points(24), MarginRight = global::OfficeIMO.OfficeLength.Points(24), MarginTop = global::OfficeIMO.OfficeLength.Points(24), MarginBottom = global::OfficeIMO.OfficeLength.Points(24), DefaultFontSize = global::OfficeIMO.OfficeLength.Points(10) })
                .Meta(title: "Workshop event badges", author: "OfficeIMO").Save(Path.Combine(folder, "example.pdf"));
        }
    }
}
