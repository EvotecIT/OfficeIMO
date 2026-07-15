namespace OfficeIMO.Pdf;

/// <summary>Configures a generated document section.</summary>
public sealed class PdfSectionOptions {
    private int _level = 1;

    /// <summary>Optional named destination; a stable document-local name is generated when omitted.</summary>
    public string? DestinationName { get; set; }

    /// <summary>Outline and TOC hierarchy level.</summary>
    public int Level {
        get => _level;
        set {
            if (value < 1 || value > 9) throw new ArgumentOutOfRangeException(nameof(value), value, "Section level must be between 1 and 9.");
            _level = value;
        }
    }

    /// <summary>Starts the section on a new page when prior page content exists.</summary>
    public bool StartOnNewPage { get; set; }
    /// <summary>Emits a heading for the section title.</summary>
    public bool IncludeHeading { get; set; } = true;
    /// <summary>Includes the section in generated tables of contents.</summary>
    public bool IncludeInTableOfContents { get; set; } = true;
    /// <summary>Optional heading style override.</summary>
    public PdfHeadingStyle? HeadingStyle { get; set; }
    /// <summary>Optional handle populated with the final output location.</summary>
    public PdfSectionReference? Reference { get; set; }

    internal PdfSectionOptions Clone(string destinationName) {
        return new PdfSectionOptions {
            DestinationName = destinationName,
            Level = Level,
            StartOnNewPage = StartOnNewPage,
            IncludeHeading = IncludeHeading,
            IncludeInTableOfContents = IncludeInTableOfContents,
            HeadingStyle = HeadingStyle?.Clone(),
            Reference = Reference
        };
    }
}
