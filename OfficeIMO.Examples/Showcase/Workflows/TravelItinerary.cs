using OfficeIMO.Pdf;

namespace OfficeIMO.Examples.Showcase.Workflows {
    /// <summary>Creates a travel handover with an itinerary, arrival instructions and clear time-zone labels.</summary>
    internal static class TravelItinerary {
        internal static void Create(string folder) {
            PdfColor muted = PdfColor.FromOfficeColor(OfficeTheme.Default.Palette.TextMuted);
            PdfDocument.Create(pdf => pdf.Content(content => content
                .Title("Workshop travel itinerary")
                .Paragraph(p => p.FontSize(9).Text("Northwind · Copenhagen · 14–15 October 2026"), defaultColor: muted)
                .PanelParagraph(p => p.Bold("All times are local to Copenhagen (CEST).\n").Text("Fictional itinerary: confirm live bookings and transport information before travelling."))
                .H2("Wednesday: arrival")
                .Table(new[] {
                    new[] { "Time", "Plan", "Location" },
                    new[] { "10:20", "Arrival and baggage collection", "Airport terminal" },
                    new[] { "11:30", "Transfer to workshop venue", "City centre" },
                    new[] { "13:00", "Working lunch and introductions", "Northwind workshop room" },
                    new[] { "14:00", "Service mapping workshop", "Northwind workshop room" }
                })
                .H2("Thursday: departure")
                .Bullets(new[] { "09:00: review the agreed actions and owners.", "12:00: leave the venue with time for the airport transfer.", "15:10: illustrative return flight departure." })
                .H2("Arrival and changes")
                .Paragraph(p => p.Text("Use the named event coordinator for venue changes. Keep booking references and personal contact details in the traveller's private copy."))
                .H2("Bring")
                .Bullets(new[] { "Travel documents and confirmed bookings", "Workshop notes and laptop charger", "The agreed accessibility and dietary arrangements" })),
                new PdfCreateOptions { ShowPageNumbers = true })
                .Meta(title: "Workshop travel itinerary", author: "OfficeIMO").Save(Path.Combine(folder, "example.pdf"));
        }
    }
}
