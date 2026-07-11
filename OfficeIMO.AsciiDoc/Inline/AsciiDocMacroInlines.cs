namespace OfficeIMO.AsciiDoc;

/// <summary>Source-backed inline macro such as <c>image:icon.svg[]</c>.</summary>
public class AsciiDocMacroInline : AsciiDocInline {
    private string _name;
    private string _target;
    private string _attributeList;

    internal AsciiDocMacroInline(AsciiDocSyntaxNode syntax, string name, string target, string attributeList) : base(syntax) {
        _name = name;
        _target = target;
        _attributeList = attributeList;
    }

    /// <summary>Macro name.</summary>
    public string Name {
        get => _name;
        set {
            string normalized = value ?? string.Empty;
            if (!AsciiDocText.IsMacroName(normalized)) throw new ArgumentException("Invalid macro name.", nameof(value));
            SetValue(ref _name, normalized);
        }
    }

    /// <summary>Macro target.</summary>
    public string Target {
        get => _target;
        set { string normalized = value ?? string.Empty; AsciiDocText.EnsureSingleLine(normalized, nameof(value)); SetValue(ref _target, normalized); }
    }

    /// <summary>Raw content between square brackets.</summary>
    public string AttributeList {
        get => _attributeList;
        set { string normalized = value ?? string.Empty; AsciiDocText.EnsureSingleLine(normalized, nameof(value)); SetValue(ref _attributeList, normalized); }
    }

    /// <summary>Parsed element attributes.</summary>
    public AsciiDocElementAttributes Attributes => AsciiDocAttributeListParser.Parse(AttributeList);

    internal override string WriteCore(AsciiDocWriterContext context) => Name + ":" + Target + "[" + AttributeList + "]";
}

/// <summary>Inline STEM expression using <c>stem:</c>, <c>latexmath:</c>, or <c>asciimath:</c>.</summary>
public sealed class AsciiDocStemInline : AsciiDocMacroInline {
    internal AsciiDocStemInline(AsciiDocSyntaxNode syntax, string name, string expression) : base(syntax, name, string.Empty, expression) { }

    /// <summary>Expression inside the macro brackets.</summary>
    public string Expression {
        get => AttributeList;
        set => AttributeList = value;
    }

    /// <summary>STEM notation selected by the macro name.</summary>
    public string Notation => string.Equals(Name, "asciimath", StringComparison.Ordinal) ? "asciimath" : "latexmath";
}

/// <summary>Inline passthrough whose content is excluded from normal substitutions.</summary>
public sealed class AsciiDocPassthroughInline : AsciiDocInline {
    private string _content;

    internal AsciiDocPassthroughInline(AsciiDocSyntaxNode syntax, string marker, string content) : base(syntax) {
        Marker = marker;
        _content = content;
    }

    /// <summary>One, two, or three plus markers.</summary>
    public string Marker { get; }

    /// <summary>Pass-through content.</summary>
    public string Content {
        get => _content;
        set => SetValue(ref _content, value ?? string.Empty);
    }

    internal override string WriteCore(AsciiDocWriterContext context) => Marker + Content + Marker;
}
