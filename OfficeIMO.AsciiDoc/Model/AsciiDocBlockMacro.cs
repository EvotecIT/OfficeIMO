namespace OfficeIMO.AsciiDoc;

/// <summary>Source-preserving AsciiDoc block macro invocation.</summary>
public sealed class AsciiDocBlockMacro : AsciiDocBlock {
    private string _name;
    private string _target;
    private string _attributeList;

    internal AsciiDocBlockMacro(
        AsciiDocSyntaxNode syntax,
        string name,
        string target,
        string attributeList,
        string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        _name = name;
        _target = target;
        _attributeList = attributeList;
    }

    /// <summary>Macro name before the <c>::</c> separator.</summary>
    public string Name {
        get => _name;
        set {
            string normalized = value ?? string.Empty;
            if (!AsciiDocText.IsMacroName(normalized)) throw new ArgumentException("Macro name contains unsupported characters.", nameof(value));
            SetValue(ref _name, normalized);
        }
    }

    /// <summary>Macro target between <c>::</c> and the attribute list.</summary>
    public string Target {
        get => _target;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            SetValue(ref _target, normalized);
        }
    }

    /// <summary>Raw attribute-list content without square brackets.</summary>
    public string AttributeList {
        get => _attributeList;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            SetValue(ref _attributeList, normalized);
        }
    }

    /// <summary>True for macro names with built-in structural meaning.</summary>
    public bool IsKnown =>
        string.Equals(Name, "include", StringComparison.Ordinal) ||
        string.Equals(Name, "image", StringComparison.Ordinal) ||
        string.Equals(Name, "video", StringComparison.Ordinal) ||
        string.Equals(Name, "audio", StringComparison.Ordinal);

    internal override string WriteCore(AsciiDocWriterContext context) =>
        Name + "::" + Target + "[" + AttributeList + "]" + EffectiveTrailingLineEnding(context);
}
