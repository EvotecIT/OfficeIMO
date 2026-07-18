namespace OfficeIMO.Email;

/// <summary>Severity of a format-native content-line validation issue.</summary>
public enum ContentLineValidationSeverity {
    /// <summary>The document can be preserved but may not interoperate as intended.</summary>
    Warning = 0,
    /// <summary>The document violates a required structural or cardinality contract.</summary>
    Error = 1
}

/// <summary>One format-native validation finding.</summary>
public sealed class ContentLineValidationIssue {
    internal ContentLineValidationIssue(string code, string message, ContentLineValidationSeverity severity,
        string? componentName = null, string? propertyName = null) {
        Code = code;
        Message = message;
        Severity = severity;
        ComponentName = componentName;
        PropertyName = propertyName;
    }

    /// <summary>Stable machine-readable issue code.</summary>
    public string Code { get; }
    /// <summary>Human-readable explanation.</summary>
    public string Message { get; }
    /// <summary>Issue severity.</summary>
    public ContentLineValidationSeverity Severity { get; }
    /// <summary>Associated component, when applicable.</summary>
    public string? ComponentName { get; }
    /// <summary>Associated property, when applicable.</summary>
    public string? PropertyName { get; }
}
