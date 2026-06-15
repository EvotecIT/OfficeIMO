namespace OfficeIMO.Rtf;

/// <summary>
/// RTF field with instruction text and rich inline result content.
/// </summary>
public sealed class RtfField : IRtfInline {
    /// <summary>Creates an RTF field.</summary>
    public RtfField(string instruction) {
        Instruction = instruction ?? throw new ArgumentNullException(nameof(instruction));
    }

    /// <summary>Raw field instruction text, for example <c>PAGE \* MERGEFORMAT</c>.</summary>
    public string Instruction { get; set; }

    /// <summary>Visible field result content.</summary>
    public RtfParagraph Result { get; } = new RtfParagraph();

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
