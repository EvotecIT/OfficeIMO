namespace OfficeIMO.Pdf;

/// <summary>
/// Configures a visual and semantic envelope around a sequence of PDF flow content.
/// </summary>
public sealed class PdfElementBuilder {
    private readonly PdfDocument _document;
    private PdfPanelStyle _style = new PdfPanelStyle {
        BorderWidth = 0D,
        PaddingX = 0D,
        PaddingY = 0D,
        SpacingBefore = 0D,
        SpacingAfter = 0D
    };
    private System.Action<PdfContentBuilder>? _content;
    private PdfSemanticRole? _semanticRole;
    private string? _alternativeText;
    private bool _hasLayoutDecoration;

    internal PdfElementBuilder(PdfDocument document) {
        _document = document;
    }

    /// <summary>Applies a reusable visual style. The style is copied when assigned.</summary>
    public PdfElementBuilder Style(PdfPanelStyle style) {
        Guard.NotNull(style, nameof(style));
        _style = style.Clone();
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Sets the content rendered inside this element.</summary>
    public PdfElementBuilder Content(System.Action<PdfContentBuilder> build) {
        Guard.NotNull(build, nameof(build));
        if (_content != null) {
            throw new System.InvalidOperationException("PDF element content can be configured only once.");
        }

        _content = build;
        return this;
    }

    /// <summary>Sets the element background fill.</summary>
    public PdfElementBuilder Background(PdfColor color) {
        _style.Background = color;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Sets a uniform border around the element.</summary>
    public PdfElementBuilder Border(PdfColor color, double width = 1D) {
        _style.BorderColor = color;
        _style.BorderWidth = width;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Sets equal padding on every side, in points.</summary>
    public PdfElementBuilder Padding(double value) => Padding(value, value);

    /// <summary>Sets vertical and horizontal padding, in points.</summary>
    public PdfElementBuilder Padding(double vertical, double horizontal) {
        _style.PaddingY = vertical;
        _style.PaddingX = horizontal;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Constrains the element to a maximum width, in points.</summary>
    public PdfElementBuilder MaxWidth(double value) {
        _style.MaxWidth = value;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Aligns a constrained element within the available content width.</summary>
    public PdfElementBuilder Align(PdfAlign align) {
        _style.Align = align;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Sets the outer vertical spacing before and after the element, in points.</summary>
    public PdfElementBuilder Spacing(double before = 0D, double after = 0D) {
        _style.SpacingBefore = before;
        _style.SpacingAfter = after;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Controls whether the entire decorated element must remain on one page.</summary>
    public PdfElementBuilder KeepTogether(bool value = true) {
        _style.KeepTogether = value;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Controls whether this element should stay with the next flow block.</summary>
    public PdfElementBuilder KeepWithNext(bool value = true) {
        _style.KeepWithNext = value;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Assigns an explicit tagged-PDF semantic role to the element.</summary>
    public PdfElementBuilder Semantic(PdfSemanticRole role, string? alternativeText = null) {
        _semanticRole = role;
        _alternativeText = alternativeText;
        return this;
    }

    internal void Commit() {
        if (_content == null) {
            throw new System.InvalidOperationException("PDF element content must be configured with Content(...).");
        }

        _document.AddElement(_content, _hasLayoutDecoration ? _style : null, _semanticRole, _alternativeText);
    }
}
