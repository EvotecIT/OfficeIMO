using System.Collections.ObjectModel;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>Standard form field kinds retained by the managed HTML render model.</summary>
public enum HtmlRenderFormFieldKind {
    /// <summary>Single-line, multiline, password, or file-select text field.</summary>
    Text,
    /// <summary>Independent check box.</summary>
    CheckBox,
    /// <summary>Single- or multi-select choice field.</summary>
    Choice,
    /// <summary>One widget belonging to a named radio-button group.</summary>
    RadioButton
}

/// <summary>
/// Backend-neutral standard HTML form metadata paired with static fallback visuals.
/// Image and SVG backends paint <see cref="Visuals"/>; interactive-capable backends may emit a native widget instead.
/// </summary>
public sealed class HtmlRenderFormField : HtmlRenderVisual {
    private readonly ReadOnlyCollection<HtmlRenderVisual> _visuals;
    private readonly ReadOnlyCollection<string> _options;
    private readonly ReadOnlyCollection<string> _optionValues;
    private readonly ReadOnlyCollection<int> _selectedOptionIndices;
    private readonly ReadOnlyCollection<string> _values;

    internal HtmlRenderFormField(
        HtmlRenderFormFieldKind fieldKind,
        string name,
        string mappingName,
        string value,
        string? placeholder,
        IEnumerable<string>? values,
        IEnumerable<string>? options,
        IEnumerable<string>? optionValues,
        IEnumerable<int>? selectedOptionIndices,
        string? radioOption,
        bool isSelected,
        bool isDisabled,
        bool isReadOnly,
        bool isRequired,
        bool isMultiline,
        bool isPassword,
        bool isFileSelect,
        bool isComboBox,
        bool allowsMultipleSelection,
        int? maximumLength,
        string? alternateName,
        OfficeFontInfo font,
        OfficeColor textColor,
        OfficeColor placeholderTextColor,
        OfficeTextAlignment textAlignment,
        OfficeColor? backgroundColor,
        OfficeColor? borderColor,
        string borderStyle,
        double borderWidth,
        double cornerRadius,
        double x,
        double y,
        double width,
        double height,
        IEnumerable<HtmlRenderVisual> visuals,
        int paintOrder,
        string? source,
        double? layoutY = null)
        : base(HtmlRenderVisualKind.FormField, x, y, width, height, paintOrder, null, source, layoutY) {
        FieldKind = fieldKind;
        Name = name ?? throw new ArgumentNullException(nameof(name));
        MappingName = mappingName ?? throw new ArgumentNullException(nameof(mappingName));
        Value = value ?? string.Empty;
        Placeholder = placeholder ?? string.Empty;
        RadioOption = radioOption;
        IsSelected = isSelected;
        IsDisabled = isDisabled;
        IsReadOnly = isReadOnly;
        IsRequired = isRequired;
        IsMultiline = isMultiline;
        IsPassword = isPassword;
        IsFileSelect = isFileSelect;
        IsComboBox = isComboBox;
        AllowsMultipleSelection = allowsMultipleSelection;
        MaximumLength = maximumLength;
        AlternateName = alternateName;
        Font = font;
        TextColor = textColor;
        PlaceholderTextColor = placeholderTextColor;
        TextAlignment = textAlignment;
        BackgroundColor = backgroundColor;
        BorderColor = borderColor;
        BorderStyle = borderStyle ?? "none";
        BorderWidth = Math.Max(0D, borderWidth);
        CornerRadius = Math.Max(0D, cornerRadius);
        _values = new List<string>(values ?? Array.Empty<string>()).AsReadOnly();
        _options = new List<string>(options ?? Array.Empty<string>()).AsReadOnly();
        _optionValues = new List<string>(optionValues ?? Array.Empty<string>()).AsReadOnly();
        _selectedOptionIndices = new List<int>(selectedOptionIndices ?? Array.Empty<int>()).AsReadOnly();
        _visuals = new List<HtmlRenderVisual>(visuals ?? throw new ArgumentNullException(nameof(visuals)))
            .OrderBy(item => item.PaintOrder)
            .ToList()
            .AsReadOnly();
    }

    /// <summary>Resolved native field kind.</summary>
    public HtmlRenderFormFieldKind FieldKind { get; }
    /// <summary>Stable field name from HTML name/id semantics or a deterministic renderer fallback.</summary>
    public string Name { get; }
    /// <summary>Original HTML field name retained for export and mapping workflows.</summary>
    public string MappingName { get; }
    /// <summary>Resolved scalar field value.</summary>
    public string Value { get; }
    /// <summary>Placeholder text shown only by the initial appearance while <see cref="Value"/> is empty.</summary>
    public string Placeholder { get; }
    /// <summary>Resolved selected values for a choice field.</summary>
    public IReadOnlyList<string> Values => _values;
    /// <summary>Resolved display options for a choice field.</summary>
    public IReadOnlyList<string> Options => _options;
    /// <summary>Export values corresponding by index to <see cref="Options"/>.</summary>
    public IReadOnlyList<string> OptionValues => _optionValues;
    /// <summary>Zero-based option identities selected by the HTML control.</summary>
    public IReadOnlyList<int> SelectedOptionIndices => _selectedOptionIndices;
    /// <summary>PDF-safe appearance-state token for a check box or radio widget.</summary>
    public string? RadioOption { get; }
    /// <summary>Whether a check box or radio widget is selected.</summary>
    public bool IsSelected { get; }
    /// <summary>Whether the HTML control is disabled and must not participate in form export.</summary>
    public bool IsDisabled { get; }
    /// <summary>Whether user editing is disabled by disabled or readonly HTML semantics.</summary>
    public bool IsReadOnly { get; }
    /// <summary>Whether HTML requires a value.</summary>
    public bool IsRequired { get; }
    /// <summary>Whether text accepts multiple lines.</summary>
    public bool IsMultiline { get; }
    /// <summary>Whether text is password content.</summary>
    public bool IsPassword { get; }
    /// <summary>Whether text represents a file selector.</summary>
    public bool IsFileSelect { get; }
    /// <summary>Whether a choice field uses a compact drop-down presentation.</summary>
    public bool IsComboBox { get; }
    /// <summary>Whether a choice field permits multiple selected options.</summary>
    public bool AllowsMultipleSelection { get; }
    /// <summary>Optional positive maximum text length.</summary>
    public int? MaximumLength { get; }
    /// <summary>Accessible field description derived from HTML/ARIA naming.</summary>
    public string? AlternateName { get; }
    /// <summary>Resolved control font.</summary>
    public OfficeFontInfo Font { get; }
    /// <summary>Resolved text color.</summary>
    public OfficeColor TextColor { get; }
    /// <summary>Resolved color used by the placeholder-only initial appearance.</summary>
    public OfficeColor PlaceholderTextColor { get; }
    /// <summary>Resolved horizontal text alignment.</summary>
    public OfficeTextAlignment TextAlignment { get; }
    /// <summary>Resolved background color.</summary>
    public OfficeColor? BackgroundColor { get; }
    /// <summary>Resolved border color.</summary>
    public OfficeColor? BorderColor { get; }
    /// <summary>Resolved uniform CSS border style retained for interactive backends.</summary>
    public string BorderStyle { get; }
    /// <summary>Resolved border width in CSS pixels.</summary>
    public double BorderWidth { get; }
    /// <summary>Resolved uniform circular corner radius in CSS pixels.</summary>
    public double CornerRadius { get; }
    /// <summary>Ordered static fallback paint used by non-interactive backends.</summary>
    public IReadOnlyList<HtmlRenderVisual> Visuals => _visuals;

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        Clone(X + offsetX, Y + offsetY, _visuals.Select((visual, index) => visual.Translate(offsetX, offsetY, index)), paintOrder, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        Clone(X + offsetX, Y + offsetY, _visuals.Select((visual, index) => visual.TranslatePaint(offsetX, offsetY, index)), paintOrder, LayoutY);

    private HtmlRenderFormField Clone(double x, double y, IEnumerable<HtmlRenderVisual> visuals, int paintOrder, double layoutY) =>
        new(FieldKind, Name, MappingName, Value, Placeholder, _values, _options, _optionValues, _selectedOptionIndices, RadioOption, IsSelected, IsDisabled, IsReadOnly, IsRequired, IsMultiline, IsPassword, IsFileSelect, IsComboBox, AllowsMultipleSelection, MaximumLength, AlternateName, Font, TextColor, PlaceholderTextColor, TextAlignment, BackgroundColor, BorderColor, BorderStyle, BorderWidth, CornerRadius, x, y, Width, Height, visuals, paintOrder, Source, layoutY);
}
