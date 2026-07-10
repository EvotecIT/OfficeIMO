namespace OfficeIMO.AsciiDoc;

/// <summary>AsciiDoc document attribute assignment or unset entry.</summary>
public sealed class AsciiDocAttributeEntry : AsciiDocBlock {
    private string _name;
    private string _value;
    private bool _isUnset;

    internal AsciiDocAttributeEntry(
        AsciiDocSyntaxNode syntax,
        string name,
        string value,
        bool isUnset,
        string trailingLineEnding)
        : base(syntax, trailingLineEnding) {
        _name = name;
        _value = value;
        _isUnset = isUnset;
    }

    /// <summary>Attribute name without surrounding colons or unset marker.</summary>
    public string Name {
        get => _name;
        set {
            string normalized = value ?? string.Empty;
            if (!AsciiDocText.IsAttributeName(normalized)) throw new ArgumentException("Attribute name contains unsupported characters.", nameof(value));
            SetValue(ref _name, normalized);
        }
    }

    /// <summary>Attribute value. Empty for boolean or unset attributes.</summary>
    public string Value {
        get => _value;
        set {
            string normalized = value ?? string.Empty;
            AsciiDocText.EnsureSingleLine(normalized, nameof(value));
            SetValue(ref _value, normalized);
        }
    }

    /// <summary>True when the entry unsets the named attribute.</summary>
    public bool IsUnset {
        get => _isUnset;
        set => SetValue(ref _isUnset, value);
    }

    internal override string WriteCore(AsciiDocWriterContext context) {
        string text = IsUnset
            ? ":" + Name + "!:"
            : ":" + Name + ":" + (Value.Length == 0 ? string.Empty : " " + Value);
        return text + EffectiveTrailingLineEnding(context);
    }
}
