namespace OfficeIMO.Pdf;

/// <summary>A compact ticket or admission-pass recipe over the normal PDF panel flow.</summary>
public sealed class PdfTicketComponent : IPdfComponent {
    private readonly System.Globalization.CultureInfo _culture;

    /// <summary>Creates a ticket recipe.</summary>
    public PdfTicketComponent(
        string title,
        string ticketNumber,
        DateTime? startsAt = null,
        string? venue = null,
        string? holder = null,
        string? instructions = null,
        System.Globalization.CultureInfo? culture = null) {
        Guard.NotNullOrWhiteSpace(title, nameof(title));
        Guard.NotNullOrWhiteSpace(ticketNumber, nameof(ticketNumber));
        Title = title;
        TicketNumber = ticketNumber;
        StartsAt = startsAt;
        Venue = venue;
        Holder = holder;
        Instructions = instructions;
        _culture = culture ?? System.Globalization.CultureInfo.InvariantCulture;
    }

    /// <summary>Event or pass title.</summary>
    public string Title { get; }
    /// <summary>Searchable ticket identifier.</summary>
    public string TicketNumber { get; }
    /// <summary>Optional start time.</summary>
    public DateTime? StartsAt { get; }
    /// <summary>Optional venue.</summary>
    public string? Venue { get; }
    /// <summary>Optional holder name.</summary>
    public string? Holder { get; }
    /// <summary>Optional instructions.</summary>
    public string? Instructions { get; }

    /// <inheritdoc />
    public void Compose(PdfItemCompose content) {
        Guard.NotNull(content, nameof(content));
        content.Panel(panel => {
            panel.H1(Title, PdfAlign.Center);
            panel.H2(TicketNumber, PdfAlign.Center);
            var details = new List<string[]>();
            if (StartsAt.HasValue) details.Add(new[] { "Starts", StartsAt.Value.ToString("g", _culture) });
            if (!string.IsNullOrWhiteSpace(Venue)) details.Add(new[] { "Venue", Venue! });
            if (!string.IsNullOrWhiteSpace(Holder)) details.Add(new[] { "Holder", Holder! });
            if (details.Count > 0) panel.Table(details, style: new PdfTableStyle { HeaderRowCount = 0, RowStripeFill = null });
            if (!string.IsNullOrWhiteSpace(Instructions)) panel.Paragraph(paragraph => paragraph.Text(Instructions!));
        }, new PanelStyle {
            BorderColor = new PdfColor(0.15, 0.15, 0.15),
            BorderWidth = 1.2,
            PaddingX = 16,
            PaddingY = 14,
            KeepTogether = true
        });
    }
}
