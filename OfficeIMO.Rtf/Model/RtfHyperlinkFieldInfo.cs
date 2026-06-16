namespace OfficeIMO.Rtf;

/// <summary>
/// Parsed semantic metadata for an RTF <c>HYPERLINK</c> field instruction.
/// The original <see cref="RtfField.Instruction"/> remains the authoritative field code.
/// </summary>
public sealed class RtfHyperlinkFieldInfo {
    /// <summary>Target URI from the first non-switch argument.</summary>
    public Uri? Target { get; set; }

    /// <summary>Optional bookmark or location switch from <c>\l</c>.</summary>
    public string? SubAddress { get; set; }

    /// <summary>Optional screen tip switch from <c>\o</c>.</summary>
    public string? ScreenTip { get; set; }

    /// <summary>Optional target frame switch from <c>\t</c>.</summary>
    public string? TargetFrame { get; set; }

    /// <summary>Optional image-map switch argument from <c>\m</c>.</summary>
    public string? ImageMap { get; set; }

    /// <summary>Creates a copy of this hyperlink field metadata.</summary>
    public RtfHyperlinkFieldInfo Clone() {
        return new RtfHyperlinkFieldInfo {
            Target = Target,
            SubAddress = SubAddress,
            ScreenTip = ScreenTip,
            TargetFrame = TargetFrame,
            ImageMap = ImageMap
        };
    }

    internal static RtfHyperlinkFieldInfo? Parse(string instruction) {
        const string hyperlinkKeyword = "HYPERLINK";
        if (!StartsWithHyperlinkKeyword(instruction, hyperlinkKeyword)) {
            return null;
        }

        int index = hyperlinkKeyword.Length;
        var info = new RtfHyperlinkFieldInfo();
        while (index < instruction.Length) {
            SkipWhiteSpace(instruction, ref index);
            if (index >= instruction.Length) {
                break;
            }

            if (instruction[index] == '\\') {
                ReadSwitch(instruction, ref index, info);
                continue;
            }

            string target = ReadToken(instruction, ref index);
            if (target.Length > 0 && Uri.TryCreate(target, UriKind.RelativeOrAbsolute, out Uri? uri)) {
                info.Target = uri;
            }
        }

        return info;
    }

    private static bool StartsWithHyperlinkKeyword(string instruction, string keyword) {
        if (!instruction.StartsWith(keyword, StringComparison.OrdinalIgnoreCase)) {
            return false;
        }

        return instruction.Length == keyword.Length || char.IsWhiteSpace(instruction[keyword.Length]);
    }

    private static void ReadSwitch(string instruction, ref int index, RtfHyperlinkFieldInfo info) {
        index++;
        int switchStart = index;
        while (index < instruction.Length && !char.IsWhiteSpace(instruction[index])) {
            index++;
        }

        string switchName = instruction.Substring(switchStart, index - switchStart);
        if (!SwitchConsumesArgument(switchName)) {
            return;
        }

        SkipWhiteSpace(instruction, ref index);
        if (index >= instruction.Length || instruction[index] == '\\') {
            return;
        }

        string value = ReadToken(instruction, ref index);
        switch (switchName.ToLowerInvariant()) {
            case "l":
                info.SubAddress = value;
                break;
            case "m":
                info.ImageMap = value;
                break;
            case "o":
                info.ScreenTip = value;
                break;
            case "t":
                info.TargetFrame = value;
                break;
        }
    }

    private static bool SwitchConsumesArgument(string switchName) {
        return string.Equals(switchName, "l", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(switchName, "m", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(switchName, "o", StringComparison.OrdinalIgnoreCase) ||
               string.Equals(switchName, "t", StringComparison.OrdinalIgnoreCase);
    }

    private static void SkipWhiteSpace(string text, ref int index) {
        while (index < text.Length && char.IsWhiteSpace(text[index])) {
            index++;
        }
    }

    private static string ReadToken(string text, ref int index) {
        return index < text.Length && text[index] == '"'
            ? ReadQuotedToken(text, ref index)
            : ReadUnquotedToken(text, ref index);
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
