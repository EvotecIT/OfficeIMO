using System;
using System.Collections.Generic;

namespace OfficeIMO.Reader;

/// <summary>
/// Source-neutral form field categories exposed by Reader adapters.
/// </summary>
public enum ReaderFormFieldKind {
    /// <summary>The source did not expose a recognized form field kind.</summary>
    Unknown,
    /// <summary>Text entry field.</summary>
    Text,
    /// <summary>Button-like field, including check boxes and radio buttons.</summary>
    Button,
    /// <summary>Choice field, including combo boxes and list boxes.</summary>
    Choice,
    /// <summary>Signature field.</summary>
    Signature
}

/// <summary>
/// Structured form field metadata extracted from a source document.
/// </summary>
public sealed class ReaderFormField {
    /// <summary>Fully qualified source field name, when available.</summary>
    public string? Name { get; set; }

    /// <summary>Partial source field name, when available.</summary>
    public string? PartialName { get; set; }

    /// <summary>User-facing alternate field name or label, when available.</summary>
    public string? AlternateName { get; set; }

    /// <summary>Export or mapping name, when available.</summary>
    public string? MappingName { get; set; }

    /// <summary>Source-specific field type token, for example Tx, Btn, Ch, or Sig for PDF.</summary>
    public string? FieldType { get; set; }

    /// <summary>Source-neutral form field kind.</summary>
    public ReaderFormFieldKind Kind { get; set; }

    /// <summary>Simple source value, when available.</summary>
    public string? Value { get; set; }

    /// <summary>Simple source values, preserving multi-select values when available.</summary>
    public IReadOnlyList<string> Values { get; set; } = Array.Empty<string>();

    /// <summary>Simple default value, when available.</summary>
    public string? DefaultValue { get; set; }

    /// <summary>Simple default values, preserving multi-select defaults when available.</summary>
    public IReadOnlyList<string> DefaultValues { get; set; } = Array.Empty<string>();

    /// <summary>Maximum text length, when the source exposes one.</summary>
    public int? MaxLength { get; set; }

    /// <summary>True when the source marks the field read-only.</summary>
    public bool IsReadOnly { get; set; }

    /// <summary>True when the source marks the field required.</summary>
    public bool IsRequired { get; set; }

    /// <summary>True when the source marks the field as excluded from export.</summary>
    public bool IsNoExport { get; set; }

    /// <summary>True when a text field allows multiple lines.</summary>
    public bool IsMultiline { get; set; }

    /// <summary>True when a text field is password-like.</summary>
    public bool IsPassword { get; set; }

    /// <summary>True when a text field uses comb formatting.</summary>
    public bool IsComb { get; set; }

    /// <summary>Number of readable choice options, when applicable.</summary>
    public int OptionCount { get; set; }

    /// <summary>Number of readable choice options selected by the current value, when applicable.</summary>
    public int SelectedOptionCount { get; set; }

    /// <summary>Number of visible source widgets associated with this field in the chunk scope.</summary>
    public int WidgetCount { get; set; }

    /// <summary>One-based page numbers where this field has widgets in the chunk scope.</summary>
    public IReadOnlyList<int> PageNumbers { get; set; } = Array.Empty<int>();

    /// <summary>Widget placement details associated with this field in the chunk scope.</summary>
    public IReadOnlyList<ReaderFormWidget> Widgets { get; set; } = Array.Empty<ReaderFormWidget>();
}

/// <summary>
/// Source-neutral form widget placement and annotation state.
/// </summary>
public sealed class ReaderFormWidget {
    /// <summary>Source field name associated with the widget, when available.</summary>
    public string? FieldName { get; set; }

    /// <summary>One-based page number containing the widget, when available.</summary>
    public int? PageNumber { get; set; }

    /// <summary>Left edge of the widget rectangle in source units.</summary>
    public double X1 { get; set; }

    /// <summary>Bottom edge of the widget rectangle in source units.</summary>
    public double Y1 { get; set; }

    /// <summary>Right edge of the widget rectangle in source units.</summary>
    public double X2 { get; set; }

    /// <summary>Top edge of the widget rectangle in source units.</summary>
    public double Y2 { get; set; }

    /// <summary>Widget rectangle width in source units.</summary>
    public double Width { get; set; }

    /// <summary>Widget rectangle height in source units.</summary>
    public double Height { get; set; }

    /// <summary>Current source appearance state name, when available.</summary>
    public string? AppearanceState { get; set; }

    /// <summary>True when the source marks the widget hidden.</summary>
    public bool IsHidden { get; set; }

    /// <summary>True when the source marks the widget printable.</summary>
    public bool IsPrint { get; set; }

    /// <summary>True when the source marks the widget read-only.</summary>
    public bool IsReadOnly { get; set; }

    /// <summary>Number of readable normal appearance states, when available.</summary>
    public int NormalAppearanceStateCount { get; set; }

    /// <summary>Readable normal appearance state names, when available.</summary>
    public IReadOnlyList<string> NormalAppearanceStates { get; set; } = Array.Empty<string>();
}
