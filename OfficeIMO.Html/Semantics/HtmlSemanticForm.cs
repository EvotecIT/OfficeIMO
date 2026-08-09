namespace OfficeIMO.Html;

/// <summary>Typed, HTML-normalized form-container state.</summary>
public sealed class HtmlSemanticForm {
    internal HtmlSemanticForm(
        string id,
        string name,
        string action,
        string method,
        string encodingType,
        string target,
        bool noValidate) {
        Id = id;
        Name = name;
        Action = action;
        Method = method;
        EncodingType = encodingType;
        Target = target;
        NoValidate = noValidate;
    }

    /// <summary>Form id used by explicitly associated controls.</summary>
    public string Id { get; }
    /// <summary>Form name.</summary>
    public string Name { get; }
    /// <summary>Authored submission action, or an empty string when none was authored.</summary>
    public string Action { get; }
    /// <summary>Effective submission method: get, post, or dialog.</summary>
    public string Method { get; }
    /// <summary>Effective submission encoding type.</summary>
    public string EncodingType { get; }
    /// <summary>Submission target browsing context.</summary>
    public string Target { get; }
    /// <summary>Whether form validation is disabled.</summary>
    public bool NoValidate { get; }
}
