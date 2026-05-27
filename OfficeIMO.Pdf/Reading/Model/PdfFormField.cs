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

    internal PdfFormField(int? objectNumber, string? name, string? partialName, string? fieldType, string? value, string? alternateName, string? mappingName, int? flags, IReadOnlyList<PdfFormWidget>? widgets = null) {
        ObjectNumber = objectNumber;
        Name = name;
        PartialName = partialName;
        FieldType = fieldType;
        Value = value;
        AlternateName = alternateName;
        MappingName = mappingName;
        Flags = flags;
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

    /// <summary>Alternate field name used as a user-facing label, when present.</summary>
    public string? AlternateName { get; }

    /// <summary>Mapping name used for export workflows, when present.</summary>
    public string? MappingName { get; }

    /// <summary>Raw field flags from /Ff, when present.</summary>
    public int? Flags { get; }

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

    private bool HasFlag(int flag) {
        return Flags.HasValue && (Flags.Value & flag) != 0;
    }
}

/// <summary>
/// Simple AcroForm widget annotation geometry read from a PDF document.
/// </summary>
public sealed class PdfFormWidget {
    internal PdfFormWidget(int? objectNumber, int? pageNumber, double x1, double y1, double x2, double y2, string? appearanceState, int? flags) {
        ObjectNumber = objectNumber;
        PageNumber = pageNumber;
        X1 = x1;
        Y1 = y1;
        X2 = x2;
        Y2 = y2;
        AppearanceState = appearanceState;
        Flags = flags;
    }

    /// <summary>Indirect object number for the widget annotation, when known.</summary>
    public int? ObjectNumber { get; }

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
}
