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
    public RtfHyperlinkFieldInfo Clone() => new RtfHyperlinkFieldInfo {
        Target = Target,
        SubAddress = SubAddress,
        ScreenTip = ScreenTip,
        TargetFrame = TargetFrame,
        ImageMap = ImageMap
    };

    /// <summary>Creates a canonical HYPERLINK field instruction with lossless quote escaping.</summary>
    public string ToInstruction() {
        var instruction = new System.Text.StringBuilder("HYPERLINK");
        AppendArgument(instruction, null, Target?.ToString());
        AppendArgument(instruction, "l", SubAddress);
        AppendArgument(instruction, "m", ImageMap);
        AppendArgument(instruction, "o", ScreenTip);
        AppendArgument(instruction, "t", TargetFrame);
        return instruction.ToString();
    }

    internal static RtfHyperlinkFieldInfo? Parse(string instruction) => Parse(RtfFieldCodeSyntax.Parse(instruction));

    internal static RtfHyperlinkFieldInfo? Parse(RtfFieldCodeSyntax syntax) {
        if (!string.Equals(syntax.Keyword, "HYPERLINK", StringComparison.OrdinalIgnoreCase)) return null;
        var info = new RtfHyperlinkFieldInfo();
        string? pendingSwitch = null;
        bool targetAssigned = false;
        foreach (RtfFieldCodeToken token in syntax.Tokens) {
            if (token.Kind is RtfFieldCodeTokenKind.Whitespace or RtfFieldCodeTokenKind.Keyword) continue;
            if (token.Kind == RtfFieldCodeTokenKind.Switch) {
                pendingSwitch = ConsumesArgument(token.Value) ? token.Value : null;
                continue;
            }
            if (pendingSwitch != null) {
                ApplySwitch(info, pendingSwitch, token.Value);
                pendingSwitch = null;
                continue;
            }
            if (!targetAssigned && Uri.TryCreate(token.Value, UriKind.RelativeOrAbsolute, out Uri? uri)) {
                info.Target = uri;
                targetAssigned = true;
            }
        }
        return info;
    }

    private static bool ConsumesArgument(string name) =>
        string.Equals(name, "l", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "m", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "o", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(name, "t", StringComparison.OrdinalIgnoreCase) ||
        name is "*" or "#" or "@";

    private static void ApplySwitch(RtfHyperlinkFieldInfo info, string name, string value) {
        switch (name.ToLowerInvariant()) {
            case "l": info.SubAddress = value; break;
            case "m": info.ImageMap = value; break;
            case "o": info.ScreenTip = value; break;
            case "t": info.TargetFrame = value; break;
        }
    }

    private static void AppendArgument(System.Text.StringBuilder instruction, string? fieldSwitch, string? value) {
        if (string.IsNullOrEmpty(value)) return;
        instruction.Append(' ');
        if (fieldSwitch != null) instruction.Append('\\').Append(fieldSwitch).Append(' ');
        instruction.Append('"')
            .Append(value!.Replace("\\", "\\\\").Replace("\"", "\\\""))
            .Append('"');
    }
}