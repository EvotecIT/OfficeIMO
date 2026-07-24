namespace OfficeIMO.Pdf;

/// <summary>
/// Common AcroForm field categories exposed by OfficeIMO.Pdf.
/// </summary>
public enum PdfFormFieldKind {
    /// <summary>The field type was not present or is not one of the common AcroForm types.</summary>
    Unknown,
    /// <summary>Text field (/FT /Tx).</summary>
    Text,
    /// <summary>Button field (/FT /Btn), including push buttons, check boxes, and radio buttons.</summary>
    Button,
    /// <summary>Choice field (/FT /Ch), including list boxes and combo boxes.</summary>
    Choice,
    /// <summary>Signature field (/FT /Sig).</summary>
    Signature
}

/// <summary>
/// Common text alignment values exposed by AcroForm /Q quadding.
/// </summary>
public enum PdfFormFieldTextAlignment {
    /// <summary>The field did not expose a recognized text alignment.</summary>
    Unknown,
    /// <summary>Left-aligned text.</summary>
    Left,
    /// <summary>Centered text.</summary>
    Center,
    /// <summary>Right-aligned text.</summary>
    Right
}

/// <summary>
/// Simple AcroForm field information read from a PDF document.
/// </summary>
public sealed class PdfFormField {
    private const int ReadOnlyFlag = 1;
    private const int RequiredFlag = 2;
    private const int NoExportFlag = 4;
    private const int MultilineFlag = 4096;
    private const int PasswordFlag = 8192;
    private const int NoToggleToOffFlag = 16384;
    private const int RadioFlag = 32768;
    private const int PushButtonFlag = 65536;
    private const int ComboFlag = 131072;
    private const int EditFlag = 262144;
    private const int SortFlag = 524288;
    private const int FileSelectFlag = 1048576;
    private const int MultiSelectFlag = 2097152;
    private const int DoNotSpellCheckFlag = 4194304;
    private const int DoNotScrollFlag = 8388608;
    private const int CombFlag = 16777216;
    private const int RichTextFlag = 33554432;
    private const int CommitOnSelectionChangeFlag = 67108864;
    private IReadOnlyList<PdfFormFieldOption>? _selectedOptions;
    private IReadOnlyList<PdfFormFieldOption>? _defaultSelectedOptions;
    private IReadOnlyList<int>? _pageNumbers;
    private IReadOnlyDictionary<int, IReadOnlyList<PdfFormWidget>>? _widgetsByPageNumber;

    internal PdfFormField(int? objectNumber, string? name, string? partialName, string? fieldType, string? value, string? alternateName, string? mappingName, int? flags, int? maxLength = null, IReadOnlyList<string>? values = null, string? defaultValue = null, IReadOnlyList<string>? defaultValues = null, string? defaultAppearance = null, int? quadding = null, IReadOnlyList<PdfFormFieldOption>? options = null, IReadOnlyList<PdfFormWidget>? widgets = null) {
        ObjectNumber = objectNumber;
        Name = name;
        PartialName = partialName;
        FieldType = fieldType;
        Value = value;
        AlternateName = alternateName;
        MappingName = mappingName;
        Flags = flags;
        MaxLength = maxLength;
        Values = values ?? Array.Empty<string>();
        DefaultValue = defaultValue;
        DefaultValues = defaultValues ?? Array.Empty<string>();
        DefaultAppearance = defaultAppearance;
        Quadding = quadding;
        Options = options ?? Array.Empty<PdfFormFieldOption>();
        Widgets = widgets ?? Array.Empty<PdfFormWidget>();
    }

    /// <summary>Indirect object number for the field dictionary, when known.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Fully qualified field name when a name can be read.</summary>
    public string? Name { get; }

    /// <summary>Partial field name from the field dictionary.</summary>
    public string? PartialName { get; }

    /// <summary>Field type name, for example Tx, Btn, Ch, or Sig, when present or inherited.</summary>
    public string? FieldType { get; }

    /// <summary>Common field kind inferred from <see cref="FieldType"/>.</summary>
    public PdfFormFieldKind Kind {
        get {
            if (string.Equals(FieldType, "Tx", StringComparison.Ordinal)) {
                return PdfFormFieldKind.Text;
            }

            if (string.Equals(FieldType, "Btn", StringComparison.Ordinal)) {
                return PdfFormFieldKind.Button;
            }

            if (string.Equals(FieldType, "Ch", StringComparison.Ordinal)) {
                return PdfFormFieldKind.Choice;
            }

            if (string.Equals(FieldType, "Sig", StringComparison.Ordinal)) {
                return PdfFormFieldKind.Signature;
            }

            return PdfFormFieldKind.Unknown;
        }
    }

    /// <summary>Simple field value formatted for wrapper display, when present.</summary>
    public string? Value { get; }

    /// <summary>Simple field values from /V, preserving array values for multi-select choice fields.</summary>
    public IReadOnlyList<string> Values { get; }

    /// <summary>Number of readable field values.</summary>
    public int ValueCount => Values.Count;

    /// <summary>True when at least one simple field value was readable.</summary>
    public bool HasValues => Values.Count > 0;

    /// <summary>Simple default field value formatted for wrapper display, when present.</summary>
    public string? DefaultValue { get; }

    /// <summary>Simple default field values from /DV, preserving array values for multi-select choice fields.</summary>
    public IReadOnlyList<string> DefaultValues { get; }

    /// <summary>Number of readable default field values.</summary>
    public int DefaultValueCount => DefaultValues.Count;

    /// <summary>True when at least one simple default field value was readable.</summary>
    public bool HasDefaultValues => DefaultValues.Count > 0;

    /// <summary>Alternate field name used as a user-facing label, when present.</summary>
    public string? AlternateName { get; }

    /// <summary>Mapping name used for export workflows, when present.</summary>
    public string? MappingName { get; }

    /// <summary>Raw field flags from /Ff, when present.</summary>
    public int? Flags { get; }

    /// <summary>Maximum text length from /MaxLen, when present on a simple field.</summary>
    public int? MaxLength { get; }

    /// <summary>Default appearance string from /DA, when present or inherited.</summary>
    public string? DefaultAppearance { get; }

    /// <summary>True when a default appearance string was readable.</summary>
    public bool HasDefaultAppearance => !string.IsNullOrEmpty(DefaultAppearance);

    /// <summary>Raw AcroForm /Q quadding value, when present or inherited.</summary>
    public int? Quadding { get; }

    /// <summary>Common text alignment inferred from /Q quadding.</summary>
    public PdfFormFieldTextAlignment TextAlignment {
        get {
            switch (Quadding) {
                case 0:
                    return PdfFormFieldTextAlignment.Left;
                case 1:
                    return PdfFormFieldTextAlignment.Center;
                case 2:
                    return PdfFormFieldTextAlignment.Right;
                default:
                    return PdfFormFieldTextAlignment.Unknown;
            }
        }
    }

    /// <summary>Choice field options from /Opt, when readable.</summary>
    public IReadOnlyList<PdfFormFieldOption> Options { get; }

    /// <summary>Number of readable choice options.</summary>
    public int OptionCount => Options.Count;

    /// <summary>True when at least one choice option was readable.</summary>
    public bool HasOptions => Options.Count > 0;

    /// <summary>Readable choice options whose export value matches the field value list.</summary>
    public IReadOnlyList<PdfFormFieldOption> SelectedOptions {
        get {
            if (_selectedOptions is not null) {
                return _selectedOptions;
            }

            _selectedOptions = GetMatchingOptions(Values, Options);
            return _selectedOptions;
        }
    }

    /// <summary>Number of readable choice options whose export value matches the field value list.</summary>
    public int SelectedOptionCount => SelectedOptions.Count;

    /// <summary>True when at least one readable choice option matches the field value list.</summary>
    public bool HasSelectedOptions => SelectedOptions.Count > 0;

    /// <summary>Readable choice options whose export value matches the default field value list.</summary>
    public IReadOnlyList<PdfFormFieldOption> DefaultSelectedOptions {
        get {
            if (_defaultSelectedOptions is not null) {
                return _defaultSelectedOptions;
            }

            _defaultSelectedOptions = GetMatchingOptions(DefaultValues, Options);
            return _defaultSelectedOptions;
        }
    }

    /// <summary>Number of readable choice options whose export value matches the default field value list.</summary>
    public int DefaultSelectedOptionCount => DefaultSelectedOptions.Count;

    /// <summary>True when at least one readable choice option matches the default field value list.</summary>
    public bool HasDefaultSelectedOptions => DefaultSelectedOptions.Count > 0;

    /// <summary>True when the common read-only field flag is set.</summary>
    public bool IsReadOnly => HasFlag(ReadOnlyFlag);

    /// <summary>True when the common required field flag is set.</summary>
    public bool IsRequired => HasFlag(RequiredFlag);

    /// <summary>True when the common no-export field flag is set.</summary>
    public bool IsNoExport => HasFlag(NoExportFlag);

    /// <summary>True when this is a text field.</summary>
    public bool IsTextField => Kind == PdfFormFieldKind.Text;

    /// <summary>True when this is a button field.</summary>
    public bool IsButtonField => Kind == PdfFormFieldKind.Button;

    /// <summary>True when this is a choice field.</summary>
    public bool IsChoiceField => Kind == PdfFormFieldKind.Choice;

    /// <summary>True when this is a signature field.</summary>
    public bool IsSignatureField => Kind == PdfFormFieldKind.Signature;

    /// <summary>True when a text field has the multiline flag set.</summary>
    public bool IsMultiline => IsTextField && HasFlag(MultilineFlag);

    /// <summary>True when a text field has the password flag set.</summary>
    public bool IsPassword => IsTextField && HasFlag(PasswordFlag);

    /// <summary>True when a text field has the file-select flag set.</summary>
    public bool IsFileSelect => IsTextField && HasFlag(FileSelectFlag);

    /// <summary>True when a text or choice field disables spell checking.</summary>
    public bool DoesNotSpellCheck => (IsTextField || IsChoiceField) && HasFlag(DoNotSpellCheckFlag);

    /// <summary>True when a text field disables scrolling.</summary>
    public bool DoesNotScroll => IsTextField && HasFlag(DoNotScrollFlag);

    /// <summary>True when a text field uses comb formatting.</summary>
    public bool IsComb => IsTextField && HasFlag(CombFlag);

    /// <summary>True when a text field has the rich-text flag set.</summary>
    public bool IsRichText => IsTextField && HasFlag(RichTextFlag);

    /// <summary>True when a button field is a push button.</summary>
    public bool IsPushButton => IsButtonField && HasFlag(PushButtonFlag);

    /// <summary>True when a button field is a radio button.</summary>
    public bool IsRadioButton => IsButtonField && HasFlag(RadioFlag);

    /// <summary>True when a button field has the no-toggle-to-off flag set.</summary>
    public bool IsNoToggleToOff => IsButtonField && HasFlag(NoToggleToOffFlag);

    /// <summary>True when a button field is a check box rather than a push button or radio button.</summary>
    public bool IsCheckBox => IsButtonField && !IsPushButton && !IsRadioButton;

    /// <summary>True when a choice field is a combo box.</summary>
    public bool IsCombo => IsChoiceField && HasFlag(ComboFlag);

    /// <summary>True when a choice combo box allows direct text editing.</summary>
    public bool IsEditableChoice => IsChoiceField && HasFlag(EditFlag);

    /// <summary>True when a choice field asks viewers to sort options.</summary>
    public bool IsSortedChoice => IsChoiceField && HasFlag(SortFlag);

    /// <summary>True when a choice field allows multiple selections.</summary>
    public bool AllowsMultipleSelection => IsChoiceField && HasFlag(MultiSelectFlag);

    /// <summary>True when a choice field commits values immediately when the selected item changes.</summary>
    public bool CommitsOnSelectionChange => IsChoiceField && HasFlag(CommitOnSelectionChangeFlag);

    /// <summary>Simple widget annotations that visually represent this field, when readable.</summary>
    public IReadOnlyList<PdfFormWidget> Widgets { get; }

    /// <summary>Number of readable widget annotations associated with this field.</summary>
    public int WidgetCount => Widgets.Count;

    /// <summary>True when at least one widget annotation was associated with this field.</summary>
    public bool HasWidgets => Widgets.Count > 0;

    /// <summary>Distinct one-based page numbers where this field has readable widget annotations.</summary>
    public IReadOnlyList<int> PageNumbers {
        get {
            if (_pageNumbers is not null) {
                return _pageNumbers;
            }

            if (Widgets.Count == 0) {
                _pageNumbers = Array.Empty<int>();
                return _pageNumbers;
            }

            var pages = new List<int>();
            var seenPages = new HashSet<int>();
            for (int i = 0; i < Widgets.Count; i++) {
                int? pageNumber = Widgets[i].PageNumber;
                if (!pageNumber.HasValue || !seenPages.Add(pageNumber.Value)) {
                    continue;
                }

                pages.Add(pageNumber.Value);
            }

            _pageNumbers = pages.Count == 0 ? Array.Empty<int>() : pages.AsReadOnly();
            return _pageNumbers;
        }
    }

    /// <summary>Number of distinct readable page numbers where this field has widgets.</summary>
    public int PageNumberCount => PageNumbers.Count;

    /// <summary>True when at least one widget for this field has a readable page number.</summary>
    public bool HasPageNumbers => PageNumberCount > 0;

    /// <summary>Readable widget annotations for this field grouped by one-based page number.</summary>
    public IReadOnlyDictionary<int, IReadOnlyList<PdfFormWidget>> WidgetsByPageNumber {
        get {
            if (_widgetsByPageNumber is not null) {
                return _widgetsByPageNumber;
            }

            var grouped = new Dictionary<int, List<PdfFormWidget>>();
            for (int i = 0; i < Widgets.Count; i++) {
                PdfFormWidget widget = Widgets[i];
                if (!widget.PageNumber.HasValue) {
                    continue;
                }

                if (!grouped.TryGetValue(widget.PageNumber.Value, out List<PdfFormWidget>? widgets)) {
                    widgets = new List<PdfFormWidget>();
                    grouped.Add(widget.PageNumber.Value, widgets);
                }

                widgets.Add(widget);
            }

            var result = new Dictionary<int, IReadOnlyList<PdfFormWidget>>();
            foreach (var item in grouped) {
                result.Add(item.Key, item.Value.AsReadOnly());
            }

            _widgetsByPageNumber = new System.Collections.ObjectModel.ReadOnlyDictionary<int, IReadOnlyList<PdfFormWidget>>(result);
            return _widgetsByPageNumber;
        }
    }

    /// <summary>Returns readable widget annotations for this field on a one-based page number.</summary>
    public IReadOnlyList<PdfFormWidget> GetWidgets(int pageNumber) {
        if (pageNumber <= 0) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), pageNumber, "Page number must be positive.");
        }

        return WidgetsByPageNumber.TryGetValue(pageNumber, out IReadOnlyList<PdfFormWidget>? widgets)
            ? widgets
            : Array.Empty<PdfFormWidget>();
    }

    private bool HasFlag(int flag) {
        return Flags.HasValue && (Flags.Value & flag) != 0;
    }

    private static IReadOnlyList<PdfFormFieldOption> GetMatchingOptions(IReadOnlyList<string> values, IReadOnlyList<PdfFormFieldOption> options) {
        if (values.Count == 0 || options.Count == 0) {
            return Array.Empty<PdfFormFieldOption>();
        }

        var selected = new List<PdfFormFieldOption>();
        var selectedValues = new HashSet<string>(values, StringComparer.Ordinal);
        for (int i = 0; i < options.Count; i++) {
            PdfFormFieldOption option = options[i];
            if (selectedValues.Contains(option.ExportValue)) {
                selected.Add(option);
            }
        }

        return selected.Count == 0 ? Array.Empty<PdfFormFieldOption>() : selected.AsReadOnly();
    }
}

/// <summary>
/// Simple AcroForm choice option read from a field /Opt array.
/// </summary>
public sealed class PdfFormFieldOption {
    internal PdfFormFieldOption(string exportValue, string displayText) {
        ExportValue = exportValue;
        DisplayText = displayText;
    }

    /// <summary>Export value used by the form field.</summary>
    public string ExportValue { get; }

    /// <summary>User-facing option text displayed by PDF viewers.</summary>
    public string DisplayText { get; }

    /// <summary>True when the option provides separate export and display text.</summary>
    public bool HasSeparateDisplayText => !string.Equals(ExportValue, DisplayText, StringComparison.Ordinal);
}

/// <summary>
/// Simple AcroForm widget annotation geometry read from a PDF document.
/// </summary>
public sealed class PdfFormWidget {
    private const int InvisibleFlag = 1;
    private const int HiddenFlag = 2;
    private const int PrintFlag = 4;
    private const int NoZoomFlag = 8;
    private const int NoRotateFlag = 16;
    private const int NoViewFlag = 32;
    private const int ReadOnlyFlag = 64;
    private const int LockedFlag = 128;
    private const int ToggleNoViewFlag = 256;
    private const int LockedContentsFlag = 512;

    internal PdfFormWidget(int? objectNumber, string? fieldName, int? pageNumber, double x1, double y1, double x2, double y2, string? appearanceState, int? flags, IReadOnlyList<string>? normalAppearanceStates = null) {
        ObjectNumber = objectNumber;
        FieldName = fieldName;
        PageNumber = pageNumber;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        AppearanceState = appearanceState;
        Flags = flags;
        NormalAppearanceStates = normalAppearanceStates ?? Array.Empty<string>();
    }

    /// <summary>Indirect object number for the widget annotation, when known.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Fully qualified form field name associated with the widget, when known.</summary>
    public string? FieldName { get; }

    /// <summary>One-based page number containing the widget annotation, when known.</summary>
    public int? PageNumber { get; }

    /// <summary>Left edge of the widget rectangle in PDF points.</summary>
    public double X1 { get; }

    /// <summary>Bottom edge of the widget rectangle in PDF points.</summary>
    public double Y1 { get; }

    /// <summary>Right edge of the widget rectangle in PDF points.</summary>
    public double X2 { get; }

    /// <summary>Top edge of the widget rectangle in PDF points.</summary>
    public double Y2 { get; }

    /// <summary>Rectangle width in PDF points.</summary>
    public double Width => X2 - X1;

    /// <summary>Rectangle height in PDF points.</summary>
    public double Height => Y2 - Y1;

    /// <summary>Current widget appearance state name from /AS, when present.</summary>
    public string? AppearanceState { get; }

    /// <summary>Raw widget annotation flags from /F, when present.</summary>
    public int? Flags { get; }

    /// <summary>True when the widget has the PDF annotation Invisible flag.</summary>
    public bool IsInvisible => HasFlag(InvisibleFlag);

    /// <summary>True when the widget has the PDF annotation Hidden flag.</summary>
    public bool IsHidden => HasFlag(HiddenFlag);

    /// <summary>True when the widget has the PDF annotation Print flag.</summary>
    public bool IsPrint => HasFlag(PrintFlag);

    /// <summary>True when the widget has the PDF annotation NoZoom flag.</summary>
    public bool IsNoZoom => HasFlag(NoZoomFlag);

    /// <summary>True when the widget has the PDF annotation NoRotate flag.</summary>
    public bool IsNoRotate => HasFlag(NoRotateFlag);

    /// <summary>True when the widget has the PDF annotation NoView flag.</summary>
    public bool IsNoView => HasFlag(NoViewFlag);

    /// <summary>True when the widget has the PDF annotation ReadOnly flag.</summary>
    public bool IsReadOnly => HasFlag(ReadOnlyFlag);

    /// <summary>True when the widget has the PDF annotation Locked flag.</summary>
    public bool IsLocked => HasFlag(LockedFlag);

    /// <summary>True when the widget has the PDF annotation ToggleNoView flag.</summary>
    public bool IsToggleNoView => HasFlag(ToggleNoViewFlag);

    /// <summary>True when the widget has the PDF annotation LockedContents flag.</summary>
    public bool IsLockedContents => HasFlag(LockedContentsFlag);

    /// <summary>Normal appearance state names from /AP /N, when the widget exposes named appearance streams.</summary>
    public IReadOnlyList<string> NormalAppearanceStates { get; }

    /// <summary>Number of readable normal appearance states.</summary>
    public int NormalAppearanceStateCount => NormalAppearanceStates.Count;

    /// <summary>True when at least one normal appearance state was readable.</summary>
    public bool HasNormalAppearanceStates => NormalAppearanceStates.Count > 0;

    /// <summary>Returns true when the widget exposes a matching normal appearance state name.</summary>
    public bool HasNormalAppearanceState(string state) {
        if (string.IsNullOrEmpty(state)) {
            return false;
        }

        for (int i = 0; i < NormalAppearanceStates.Count; i++) {
            if (string.Equals(NormalAppearanceStates[i], state, StringComparison.Ordinal)) {
                return true;
            }
        }

        return false;
    }

    private bool HasFlag(int flag) {
        return Flags.HasValue && (Flags.Value & flag) != 0;
    }
}
