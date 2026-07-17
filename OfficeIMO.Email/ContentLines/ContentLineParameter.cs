namespace OfficeIMO.Email;

/// <summary>One named parameter attached to an iCalendar or vCard content line.</summary>
public sealed class ContentLineParameter {
    private readonly List<string> _values = new List<string>();

    /// <summary>Creates an empty parameter with the supplied name.</summary>
    public ContentLineParameter(string name) {
        Name = ContentLineSyntax.RequireToken(name, nameof(name));
    }

    /// <summary>Creates a parameter with one or more values.</summary>
    public ContentLineParameter(string name, params string[] values) : this(name) {
        if (values == null) throw new ArgumentNullException(nameof(values));
        foreach (string value in values) _values.Add(value ?? throw new ArgumentNullException(nameof(values)));
    }

    /// <summary>Parameter name. Names are compared case-insensitively by lookup helpers.</summary>
    public string Name { get; set; }

    /// <summary>Ordered parameter values.</summary>
    public IList<string> Values => _values;
}
