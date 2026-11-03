using OfficeIMO.Drawing;
using System.Collections.Generic;

namespace OfficeIMO.OneNote.Html;

/// <summary>Controls visual-preserving HTML generated from the native OneNote page canvas.</summary>
public sealed class OneNoteVisualHtmlExportOptions {
    /// <summary>Page rendering options shared with image and PDF export.</summary>
    public OneNotePageRenderingOptions PageRendering { get; set; } = new OneNotePageRenderingOptions();

    /// <summary>Standalone document title. The notebook or section name is used when absent.</summary>
    public string? DocumentTitle { get; set; }

    /// <summary>BCP 47 document language written to standalone HTML.</summary>
    public string Language { get; set; } = "en";

    /// <summary>Whether encoded semantic page text is included for assistive technology and indexing.</summary>
    public bool IncludeAccessibleText { get; set; } = true;

    /// <summary>Whether the built-in responsive canvas stylesheet is included.</summary>
    public bool IncludeDefaultStyles { get; set; } = true;

    /// <summary>Optional caller-owned collection that receives page-mapping and image-fallback diagnostics.</summary>
    public ICollection<OfficeImageExportDiagnostic>? DiagnosticSink { get; set; }

    /// <summary>Produces an embeddable SVG-page fragment instead of a standalone HTML5 document.</summary>
    public bool Fragment { get; set; }

    /// <summary>Throws <see cref="OfficeConversionException"/> when the conversion reports loss.</summary>
    public bool FailOnLoss { get; set; }

    /// <summary>Creates an independent copy of these settings; the diagnostic sink is shared.</summary>
    public OneNoteVisualHtmlExportOptions Clone() => new OneNoteVisualHtmlExportOptions {
        PageRendering = PageRendering?.Clone() ?? new OneNotePageRenderingOptions(),
        DocumentTitle = DocumentTitle,
        Language = Language,
        IncludeAccessibleText = IncludeAccessibleText,
        IncludeDefaultStyles = IncludeDefaultStyles,
        DiagnosticSink = DiagnosticSink,
        Fragment = Fragment,
        FailOnLoss = FailOnLoss
    };

    internal void Validate() {
        if (PageRendering == null) throw new InvalidOperationException("Page rendering options cannot be null.");
        if (string.IsNullOrWhiteSpace(Language)) throw new ArgumentException("HTML document language cannot be empty.", nameof(Language));
    }
}
