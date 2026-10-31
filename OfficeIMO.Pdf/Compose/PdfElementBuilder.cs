namespace OfficeIMO.Pdf;

/// <summary>
/// Configures a visual and semantic envelope around a sequence of PDF flow content.
/// </summary>
public sealed class PdfElementBuilder {
    private readonly PdfDocument _document;
    private PdfPanelStyle _style = new PdfPanelStyle {
        BorderWidth = global::OfficeIMO.OfficeLength.Points(0D),
        PaddingX = global::OfficeIMO.OfficeLength.Points(0D),
        PaddingY = global::OfficeIMO.OfficeLength.Points(0D),
        SpacingBefore = global::OfficeIMO.OfficeLength.Points(0D),
        SpacingAfter = global::OfficeIMO.OfficeLength.Points(0D)
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
    public PdfElementBuilder Background(global::OfficeIMO.OfficeColor color) {
        _style.Background = color;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Sets a uniform border around the element.</summary>
    public PdfElementBuilder Border(global::OfficeIMO.OfficeColor color, global::OfficeIMO.OfficeLength? width = null) {
        _style.BorderColor = color;
        _style.BorderWidth = (width ?? global::OfficeIMO.OfficeLength.Points(1D));
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Sets equal padding on every side, in points.</summary>
    public PdfElementBuilder Padding(global::OfficeIMO.OfficeLength value) => Padding(value, value);

    /// <summary>Sets vertical and horizontal padding, in points.</summary>
    public PdfElementBuilder Padding(global::OfficeIMO.OfficeLength vertical, global::OfficeIMO.OfficeLength horizontal) {
        _style.PaddingY = vertical;
        _style.PaddingX = horizontal;
        _hasLayoutDecoration = true;
        return this;
    }

    /// <summary>Constrains the element to a maximum width, in points.</summary>
    public PdfElementBuilder MaxWidth(global::OfficeIMO.OfficeLength value) {
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
    public PdfElementBuilder Spacing(global::OfficeIMO.OfficeLength before = default, global::OfficeIMO.OfficeLength after = default) {
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
