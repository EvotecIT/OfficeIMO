using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    /// <summary>Gets or sets directly authored page-break-before formatting. Null restores style inheritance; false overrides an enabled style.</summary>
    public bool? PageBreakBeforeOverride {
        get => ReadFormattingOverride<PageBreakBefore>();
        set => WriteFormattingOverride<PageBreakBefore>(value);
    }

    /// <summary>Gets or sets directly authored keep-with-next formatting. Null restores style inheritance; false overrides an enabled style.</summary>
    public bool? KeepWithNextOverride {
        get => ReadFormattingOverride<KeepNext>();
        set => WriteFormattingOverride<KeepNext>(value);
    }

    /// <summary>Gets or sets directly authored keep-lines-together formatting. Null restores style inheritance; false overrides an enabled style.</summary>
    public bool? KeepLinesTogetherOverride {
        get => ReadFormattingOverride<KeepLines>();
        set => WriteFormattingOverride<KeepLines>(value);
    }

    /// <summary>Gets or sets directly authored widow/orphan control. Null restores style inheritance; false disables the control.</summary>
    public bool? AvoidWidowAndOrphanOverride {
        get => ReadFormattingOverride<WidowControl>();
        set => WriteFormattingOverride<WidowControl>(value);
    }

    /// <summary>Gets or sets suppression of spacing between paragraphs of the same style. Null restores style inheritance.</summary>
    public bool? ContextualSpacing {
        get => ReadFormattingOverride<DocumentFormat.OpenXml.Wordprocessing.ContextualSpacing>();
        set => WriteFormattingOverride<DocumentFormat.OpenXml.Wordprocessing.ContextualSpacing>(value);
    }

    /// <summary>Gets or sets paragraph line-number suppression. Null restores style inheritance. This preserves the Word setting and does not enable line-number rendering in PDF.</summary>
    public bool? SuppressLineNumbers {
        get => ReadFormattingOverride<DocumentFormat.OpenXml.Wordprocessing.SuppressLineNumbers>();
        set => WriteFormattingOverride<DocumentFormat.OpenXml.Wordprocessing.SuppressLineNumbers>(value);
    }

    /// <summary>Gets or sets paragraph automatic-hyphenation suppression. Null restores style inheritance. Automatic Word hyphenation is not supplied by this property.</summary>
    public bool? SuppressAutoHyphens {
        get => ReadFormattingOverride<DocumentFormat.OpenXml.Wordprocessing.SuppressAutoHyphens>();
        set => WriteFormattingOverride<DocumentFormat.OpenXml.Wordprocessing.SuppressAutoHyphens>(value);
    }

    /// <summary>Gets or sets use of left/right indents as inside/outside indents. Null restores style inheritance. Fixed-layout exporters have their own mirrored-layout support limits.</summary>
    public bool? MirrorIndents {
        get => ReadFormattingOverride<DocumentFormat.OpenXml.Wordprocessing.MirrorIndents>();
        set => WriteFormattingOverride<DocumentFormat.OpenXml.Wordprocessing.MirrorIndents>(value);
    }

    /// <summary>Gets or sets the directly authored outline level: 0-8 for heading levels 1-9, 9 for body text, or null to inherit. Word uses the built-in heading level for Heading1-Heading9 styles regardless of this authored value.</summary>
    public int? OutlineLevel {
        get => _paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value;
        set {
            if (value is < 0 or > 9) throw new ArgumentOutOfRangeException(nameof(value), "Outline level must be between 0 and 9.");
            if (value == null) {
                _paragraph.ParagraphProperties?.RemoveAllChildren<DocumentFormat.OpenXml.Wordprocessing.OutlineLevel>();
            } else {
                (_paragraph.ParagraphProperties ??= new ParagraphProperties()).OutlineLevel =
                    new DocumentFormat.OpenXml.Wordprocessing.OutlineLevel { Val = value.Value };
            }
        }
    }

    private bool? ReadFormattingOverride<T>() where T : OnOffType {
        T? element = _paragraph.ParagraphProperties?.GetFirstChild<T>();
        return element == null ? null : element.Val?.Value ?? true;
    }

    private void WriteFormattingOverride<T>(bool? value) where T : OnOffType, new() {
        ParagraphProperties? properties = _paragraph.ParagraphProperties;
        properties?.RemoveAllChildren<T>();
        if (value.HasValue) {
            properties ??= _paragraph.ParagraphProperties = new ParagraphProperties();
            properties.AddChild(new T { Val = value.Value }, true);
        }
    }
}
