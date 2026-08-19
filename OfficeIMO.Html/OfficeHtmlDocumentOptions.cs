namespace OfficeIMO.Html;

/// <summary>
/// Options used by shared OfficeIMO HTML document shell helpers.
/// </summary>
public sealed class OfficeHtmlDocumentOptions {
    /// <summary>When true, emits a complete HTML document; otherwise emits only the supplied fragment.</summary>
    public bool EmitDocumentShell { get; set; } = true;

    /// <summary>HTML document title.</summary>
    public string? Title { get; set; } = "OfficeIMO HTML";

    /// <summary>BCP 47 language tag assigned to the generated document element.</summary>
    public string? Language { get; set; } = "en";

    /// <summary>Theme applied when default shell styles are included.</summary>
    public OfficeVisualThemeKind Theme { get; set; } = OfficeVisualThemeKind.WordLike;

    /// <summary>When true, emits the shared OfficeIMO CSS shell.</summary>
    public bool IncludeDefaultStyles { get; set; } = true;

    /// <summary>Optional CSS class assigned to the generated body.</summary>
    public string BodyClass { get; set; } = "officeimo-html";

    /// <summary>Line ending used by generated HTML.</summary>
    public string NewLine { get; set; } = "\n";

    /// <summary>Creates an independent copy suitable for one conversion.</summary>
    public OfficeHtmlDocumentOptions Clone() => new() {
        EmitDocumentShell = EmitDocumentShell,
        Title = Title,
        Language = Language,
        Theme = Theme,
        IncludeDefaultStyles = IncludeDefaultStyles,
        BodyClass = BodyClass,
        NewLine = NewLine
    };

    /// <summary>Validates the bounded document-output contract.</summary>
    public void Validate() {
        if (IncludeDefaultStyles && !Enum.IsDefined(typeof(OfficeVisualThemeKind), Theme)) {
            throw new ArgumentOutOfRangeException(nameof(Theme), Theme, "Office HTML theme is not supported.");
        }
        if (Language != null && string.IsNullOrWhiteSpace(Language)) {
            throw new ArgumentException("HTML document language cannot be empty.", nameof(Language));
        }
        if (NewLine != "\n" && NewLine != "\r\n" && NewLine != "\r") {
            throw new ArgumentException("HTML document newline must be LF, CRLF, or CR.", nameof(NewLine));
        }
    }
}
