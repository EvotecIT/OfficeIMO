namespace OfficeIMO.Yaml;

/// <summary>How <see cref="YamlFormatter"/> indents block sequences under a mapping key.</summary>
public enum YamlSequenceIndentation {
    /// <summary>The style the document uses most: what saving new nodes would follow.</summary>
    Preserve,
    /// <summary>The dash level with the key: <c>key:\n- item</c>.</summary>
    Compact,
    /// <summary>The dash indented under the key by the indent width: <c>key:\n  - item</c>, as Kubernetes manifests and Helm charts write it.</summary>
    Indented
}

/// <summary>Which quotes <see cref="YamlFormatter"/> gives quoted strings.</summary>
public enum YamlQuoteStyle {
    /// <summary>Each scalar keeps its own quotes.</summary>
    Preserve,
    /// <summary>Single quotes, except where a value needs double-quoted escapes.</summary>
    Single,
    /// <summary>Double quotes.</summary>
    Double
}

/// <summary>The line break <see cref="YamlFormatter"/> writes.</summary>
public enum YamlLineEnding {
    /// <summary>The break most lines of the input use.</summary>
    Preserve,
    /// <summary><c>\n</c>.</summary>
    Lf,
    /// <summary><c>\r\n</c>.</summary>
    CrLf
}

/// <summary>
/// Layout settings for <see cref="YamlFormatter"/>. Every setting left at its default follows the document's own
/// conventions, so formatting with default options only removes inconsistencies.
/// </summary>
public sealed class YamlFormatOptions {
    private int? _indent;

    /// <summary>
    /// Gets or sets how far nested mapping keys are indented from their parent key, or <see langword="null"/> (the
    /// default) for the width the document uses. 1 to 9.
    /// </summary>
    public int? IndentWidth {
        get => _indent;
        set {
            if (value is < 1 or > 9) throw new ArgumentOutOfRangeException(nameof(value), value, "IndentWidth must be between 1 and 9.");
            _indent = value;
        }
    }

    /// <summary>Gets or sets how sequences under a key are indented. Defaults to <see cref="YamlSequenceIndentation.Preserve"/>.</summary>
    public YamlSequenceIndentation SequenceIndentation { get; set; }

    /// <summary>
    /// Gets or sets the quotes of quoted strings. Defaults to <see cref="YamlQuoteStyle.Preserve"/>. Plain scalars stay
    /// plain, and strings that would read as another type stay quoted.
    /// </summary>
    public YamlQuoteStyle Quotes { get; set; }

    /// <summary>Gets or sets the line break. Defaults to <see cref="YamlLineEnding.Preserve"/>.</summary>
    public YamlLineEnding LineEnding { get; set; }

    /// <summary>
    /// Gets or sets whether the text ends with a line break. Defaults to <see langword="true"/>. With
    /// <see langword="false"/> the final break is left out unless a block scalar at the end needs it for its value.
    /// </summary>
    public bool FinalNewline { get; set; } = true;

    internal YamlWriterOptions ToWriterOptions(YamlWriterOptions inferred) {
        int indent = _indent ?? inferred.MappingIndent;
        var options = new YamlWriterOptions { MappingIndent = indent };
        switch (SequenceIndentation) {
            case YamlSequenceIndentation.Compact:
                options.SequenceIndent = 2;
                break;
            case YamlSequenceIndentation.Indented:
                SetIndented(options, indent);
                break;
            default:
                // An explicit width scales the document's style: offset 0 stays compact, any offset becomes the width.
                bool indented = inferred.SequenceDashOffset > 0;
                if (_indent is null) {
                    options.SequenceIndent = inferred.SequenceIndent;
                    options.SequenceDashOffset = inferred.SequenceDashOffset;
                } else if (indented) {
                    SetIndented(options, indent);
                }
                break;
        }
        options.NewLine = LineEnding switch {
            YamlLineEnding.Lf => "\n",
            YamlLineEnding.CrLf => "\r\n",
            _ => inferred.NewLine
        };
        return options;
    }

    /// <summary>The dash at the indent width, and the content after it; the writer allows dash offsets up to 8.</summary>
    private static void SetIndented(YamlWriterOptions options, int indent) {
        int offset = Math.Min(indent, 8);
        options.SequenceIndent = offset + 2;
        options.SequenceDashOffset = offset;
    }

    internal YamlStyle QuoteStyle => Quotes switch {
        YamlQuoteStyle.Single => YamlStyle.SingleQuoted,
        YamlQuoteStyle.Double => YamlStyle.DoubleQuoted,
        _ => YamlStyle.None
    };
}
