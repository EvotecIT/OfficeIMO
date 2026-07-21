namespace OfficeIMO.Pdf;

/// <summary>One semantic section in a reusable report component.</summary>
public sealed class PdfReportSection {
    /// <summary>Creates a report section from a heading, body, and optional bullet list.</summary>
    public PdfReportSection(string title, string? body = null, IEnumerable<string>? bullets = null) {
        Guard.NotNullOrWhiteSpace(title, nameof(title));
        Title = title;
        Body = body;
        Bullets = Snapshot(bullets);
    }

    /// <summary>Section heading.</summary>
    public string Title { get; }
    /// <summary>Optional narrative body.</summary>
    public string? Body { get; }
    /// <summary>Optional bullet items.</summary>
    public IReadOnlyList<string> Bullets { get; }

    private static IReadOnlyList<string> Snapshot(IEnumerable<string>? values) {
        if (values == null) return Array.Empty<string>();
        var result = new List<string>();
        foreach (string value in values) {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Report bullets cannot be empty.", nameof(values));
            result.Add(value);
        }
        return result.AsReadOnly();
    }
}

/// <summary>A small reusable report recipe implemented entirely over the canonical PDF flow engine.</summary>
public sealed class PdfReportComponent : IPdfComponent {
    private readonly KeyValuePair<string, string?>[] _facts;
    private readonly PdfReportSection[] _sections;

    /// <summary>Creates a report recipe.</summary>
    public PdfReportComponent(
        string title,
        string? summary = null,
        IEnumerable<KeyValuePair<string, string?>>? facts = null,
        IEnumerable<PdfReportSection>? sections = null) {
        Guard.NotNullOrWhiteSpace(title, nameof(title));
        Title = title;
        Summary = summary;
        _facts = facts?.ToArray() ?? Array.Empty<KeyValuePair<string, string?>>();
        _sections = sections?.ToArray() ?? Array.Empty<PdfReportSection>();
        if (_facts.Any(static fact => string.IsNullOrWhiteSpace(fact.Key))) throw new ArgumentException("Report fact names cannot be empty.", nameof(facts));
        if (_sections.Any(static section => section == null)) throw new ArgumentException("Report sections cannot contain null entries.", nameof(sections));
    }

    /// <summary>Report title.</summary>
    public string Title { get; }
    /// <summary>Optional executive summary.</summary>
    public string? Summary { get; }

    /// <inheritdoc />
    public void Compose(PdfItemCompose content) {
        Guard.NotNull(content, nameof(content));
        content.H1(Title).HR();
        if (!string.IsNullOrWhiteSpace(Summary)) content.Paragraph(paragraph => paragraph.Text(Summary!));
        if (_facts.Length > 0) {
            content.Table(
                _facts.Select(static fact => new[] { fact.Key, fact.Value ?? string.Empty }),
                style: new PdfTableStyle { HeaderRowCount = 0, RowStripeFill = null, SpacingAfter = 10 });
        }
        foreach (PdfReportSection section in _sections) {
            content.H2(section.Title);
            if (!string.IsNullOrWhiteSpace(section.Body)) content.Paragraph(paragraph => paragraph.Text(section.Body!));
            if (section.Bullets.Count > 0) content.Bullets(section.Bullets);
        }
    }
}
