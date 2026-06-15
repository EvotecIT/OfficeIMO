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
            Hyperlink = TryParseHyperlinkTarget(_instruction);
        }
    }

    /// <summary>Visible field result content.</summary>
    public RtfParagraph Result { get; } = new RtfParagraph();

    /// <summary>Parsed hyperlink target for <c>HYPERLINK</c> fields. The raw <see cref="Instruction"/> remains authoritative.</summary>
    public Uri? Hyperlink { get; set; }

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

    private static Uri? TryParseHyperlinkTarget(string instruction) {
        const string hyperlinkKeyword = "HYPERLINK";
        if (!instruction.StartsWith(hyperlinkKeyword, StringComparison.OrdinalIgnoreCase)) {
            return null;
        }

        if (instruction.Length > hyperlinkKeyword.Length && !char.IsWhiteSpace(instruction[hyperlinkKeyword.Length])) {
            return null;
        }

        int index = hyperlinkKeyword.Length;
        while (index < instruction.Length) {
            SkipWhiteSpace(instruction, ref index);
            if (index >= instruction.Length) {
                return null;
            }

            if (instruction[index] == '\\') {
                SkipSwitch(instruction, ref index);
                continue;
            }

            string target = instruction[index] == '"'
                ? ReadQuotedToken(instruction, ref index)
                : ReadUnquotedToken(instruction, ref index);
            return Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out Uri? uri) ? uri : null;
        }

        return null;
    }

    private static void SkipWhiteSpace(string text, ref int index) {
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }
    }

    private static void SkipSwitch(string text, ref int index) {
        index++;
        int switchStart = index;
        while (index < text.Length && !char.IsWhiteSpace(text[index])) {
            index++;
        }

        string switchName = text.Substring(switchStart, index - switchStart);
        if (!SwitchConsumesArgument(switchName)) {
            return;
        }

        SkipWhiteSpace(text, ref index);
        if (index < text.Length && text[index] == '"') {
            ReadQuotedToken(text, ref index);
            return;
        }

        while (index < text.Length && !char.IsWhiteSpace(text[index]) && text[index] != '\\') {
            index++;
        }
    }

    private static bool SwitchConsumesArgument(string switchName) {
        return string.Equals(switchName, "l", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(switchName, "m", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(switchName, "o", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(switchName, "t", StringComparison.OrdinalIgnoreCase);
    }

    private static string ReadQuotedToken(string text, ref int index) {
        index++;
        var builder = new System.Text.StringBuilder();
        while (index < text.Length) {
            char value = text[index++];
            if (value == '"') {
                break;
            }

            builder.Append(value);
        }

        return builder.ToString();
    }

    private static string ReadUnquotedToken(string text, ref int index) {
        int start = index;
        while (index < text.Length && !char.IsWhiteSpace(text[index])) {
            index++;
        }

        return text.Substring(start, index - start);
    }
}
