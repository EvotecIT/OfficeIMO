namespace OfficeIMO.OpenDocument;

internal static class OdfImageMirrorCodec {
    internal static OdfImageMirror? Parse(string? lexical) {
        if (lexical == null) return null;
        string[] tokens = lexical.Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
        if (tokens.Length == 1 && tokens[0] == "none") return OdfImageMirror.None;
        OdfImageMirror mirror = OdfImageMirror.None;
        if (tokens.Length is < 1 or > 2) throw InvalidSyntax();
        foreach (string token in tokens) {
            OdfImageMirror flag = token switch {
                "horizontal" => OdfImageMirror.Horizontal,
                "vertical" => OdfImageMirror.Vertical,
                "horizontal-on-odd" => OdfImageMirror.HorizontalOnOddPages,
                "horizontal-on-even" => OdfImageMirror.HorizontalOnEvenPages,
                _ => throw InvalidSyntax()
            };
            if ((mirror & flag) != 0) throw InvalidSyntax();
            mirror |= flag;
        }
        if (!IsValid(mirror)) throw InvalidSyntax();
        return mirror;
    }

    internal static string? Serialize(OdfImageMirror? mirror) {
        if (!mirror.HasValue) return null;
        OdfImageMirror value = mirror.Value;
        if (!IsValid(value)) throw new ArgumentOutOfRangeException(nameof(mirror), "Use at most one horizontal mirror mode, optionally combined with vertical mirroring.");
        if (value == OdfImageMirror.None) return "none";
        string horizontal = (value & ~OdfImageMirror.Vertical) switch {
            OdfImageMirror.Horizontal => "horizontal",
            OdfImageMirror.HorizontalOnOddPages => "horizontal-on-odd",
            OdfImageMirror.HorizontalOnEvenPages => "horizontal-on-even",
            _ => string.Empty
        };
        return (value & OdfImageMirror.Vertical) != 0 ? (horizontal.Length == 0 ? "vertical" : horizontal + " vertical") : horizontal;
    }

    private static bool IsValid(OdfImageMirror mirror) => mirror is OdfImageMirror.None or OdfImageMirror.Vertical ||
        (mirror & ~OdfImageMirror.Vertical) is OdfImageMirror.Horizontal or OdfImageMirror.HorizontalOnOddPages or OdfImageMirror.HorizontalOnEvenPages;
    private static InvalidDataException InvalidSyntax() => new InvalidDataException("Image mirroring must be none, vertical, or one horizontal mode optionally combined with vertical.");
}
