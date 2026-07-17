namespace OfficeIMO.Email;

/// <summary>One ordered property in an iCalendar component or vCard.</summary>
public sealed class ContentLineProperty {
    private readonly List<ContentLineParameter> _parameters = new List<ContentLineParameter>();

    /// <summary>Creates a property with its raw, still format-escaped value.</summary>
    public ContentLineProperty(string name, string value = "") {
        Name = ContentLineSyntax.RequireToken(name, nameof(name));
        Value = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Optional vCard property group.</summary>
    public string? Group { get; set; }

    /// <summary>Property name. Names are compared case-insensitively by lookup helpers.</summary>
    public string Name { get; set; }

    /// <summary>Ordered parameters, including repeated parameter names.</summary>
    public IList<ContentLineParameter> Parameters => _parameters;

    /// <summary>Raw property value after the first unquoted colon.</summary>
    public string Value { get; set; }

    /// <summary>Returns the first parameter matching <paramref name="name"/>.</summary>
    public ContentLineParameter? GetParameter(string name) => _parameters.FirstOrDefault(parameter =>
        string.Equals(parameter.Name, name, StringComparison.OrdinalIgnoreCase));

    /// <summary>Replaces every matching parameter with one parameter containing the supplied values.</summary>
    public ContentLineProperty SetParameter(string name, params string[] values) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        for (int index = _parameters.Count - 1; index >= 0; index--) {
            if (string.Equals(_parameters[index].Name, name, StringComparison.OrdinalIgnoreCase))
                _parameters.RemoveAt(index);
        }
        _parameters.Add(new ContentLineParameter(name, values));
        return this;
    }
}
