using System.IO;
using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows;

/// <summary>Creates a two-page workshop pack with a shared header and numbered footer.</summary>
internal static class WorkshopPack {
    internal static void Create(string folder) {
        OfficeColor rule = OfficeTheme.Default.TableStyle.BorderColor;
        var options = new PdfCreateOptions {
            PageSetup = {
                Header = { Show = true, Format = "Northwind · Service design", FontSize = OfficeLength.Points(8.5), Alignment = OfficeAlignment.Left, Color = OfficeTheme.Default.Palette.TextMuted },
                Footer = { Show = true, FontSize = OfficeLength.Points(8.5), Alignment = OfficeAlignment.Right }
            }
        };
        PdfDocument.Create(pdf => pdf.Content(content => content
            .Title("Service design workshop")
            .Subtitle("A half-day session to turn service feedback into a delivery plan.")
            .H2("What we will leave with")
            .AddList(new[] { "One agreed service journey", "Three prioritized improvements", "An owner and next step for each improvement" })
            .H2("Agenda")
            .AddTable(new[] {
                new[] { "Time", "Session", "Output" },
                new[] { "09:00", "Current journey", "Shared map" },
                new[] { "09:45", "Friction and evidence", "Ranked pain points" },
                new[] { "10:45", "Options and trade-offs", "Three candidates" },
                new[] { "11:30", "Commitments", "Action list" }
            })
            .PageBreak()
            .H1("Exercise sheet")
            .H3("1. Describe one difficult moment")
            .AddParagraph(p => p.AddText("Who is trying to do what, and where do they get stuck? Use an observed example."))
            .Spacer(OfficeLength.Points(130))
            .HR(OfficeLength.Points(0.5), rule, OfficeLength.Points(0), OfficeLength.Points(18))
            .H3("2. Propose the smallest useful change")
            .AddParagraph(p => p.AddText("Describe what the person will be able to do after the change."))
            .Spacer(OfficeLength.Points(130))
            .HR(OfficeLength.Points(0.5), rule, OfficeLength.Points(0), OfficeLength.Points(18))
            .H3("3. Choose the first check")
            .AddParagraph(p => p.AddText("What evidence would show that the change helped? Who will collect it?"))), options)
            .Meta(title: "Service design workshop pack", author: "OfficeIMO")
            .Save(Path.Combine(folder, "example.pdf"));
    }
}
