namespace OfficeIMO.Rtf;

/// <summary>
/// RTF field with instruction text and rich inline result content.
/// </summary>
public sealed class RtfField : IRtfInline {
    private string _instruction = string.Empty;

    /// <summary>Creates an RTF field.</summary>
    public RtfField(string instruction) {
        Instruction = instruction;
    }

    /// <summary>Raw field instruction text, for example <c>PAGE \* MERGEFORMAT</c>.</summary>
    public string Instruction {
        get => _instruction;
        set {
            _instruction = value ?? throw new ArgumentNullException(nameof(Instruction));
            HyperlinkField = RtfHyperlinkFieldInfo.Parse(_instruction);
        }
    }

    /// <summary>Visible field result content.</summary>
    public RtfParagraph Result { get; } = new RtfParagraph();

    /// <summary>Parsed metadata for <c>HYPERLINK</c> fields. The raw <see cref="Instruction"/> remains authoritative.</summary>
    public RtfHyperlinkFieldInfo? HyperlinkField { get; set; }

    /// <summary>Gets whether this is a Word <c>EQ</c> mathematical equation field.</summary>
    public bool IsEquation {
        get {
            string trimmed = Instruction.TrimStart();
            return trimmed.StartsWith("EQ", StringComparison.OrdinalIgnoreCase)
                && (trimmed.Length == 2 || char.IsWhiteSpace(trimmed[2]));
        }
    }

    /// <summary>Parsed hyperlink target for <c>HYPERLINK</c> fields.</summary>
    public Uri? Hyperlink {
        get => HyperlinkField?.Target;
        set {
            if (value == null) {
                if (HyperlinkField != null) {
                    HyperlinkField.Target = null;
                }

                return;
            }

            HyperlinkField ??= new RtfHyperlinkFieldInfo();
            HyperlinkField.Target = value;
        }
    }

    /// <summary>Creates hyperlink field metadata if it does not already exist and returns it.</summary>
    public RtfHyperlinkFieldInfo GetOrCreateHyperlinkField() {
        HyperlinkField ??= new RtfHyperlinkFieldInfo();
        return HyperlinkField;
    }

    /// <summary>Optional Word form-field metadata from an RTF <c>\ffdata</c> destination.</summary>
    public RtfFormFieldData? FormFieldData { get; set; }

    /// <summary>Creates form-field metadata if it does not already exist and returns it.</summary>
    public RtfFormFieldData GetOrCreateFormFieldData() {
        FormFieldData ??= new RtfFormFieldData();
        return FormFieldData;
    }

    /// <summary>Configures Word form-field metadata for this field.</summary>
    public RtfField SetFormFieldData(Action<RtfFormFieldData> configure) {
        if (configure == null) throw new ArgumentNullException(nameof(configure));
        configure(GetOrCreateFormFieldData());
        return this;
    }

    /// <summary>Adds visible result text.</summary>
    public RtfRun AddText(string text) {
        return Result.AddText(text);
    }

    /// <summary>Returns visible field result text.</summary>
    public string ToPlainText() {
        return Result.ToPlainText();
    }

}
