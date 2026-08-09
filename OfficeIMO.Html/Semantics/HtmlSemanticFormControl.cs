namespace OfficeIMO.Html;

/// <summary>Typed, HTML-normalized form-control state.</summary>
public sealed class HtmlSemanticFormControl {
    internal HtmlSemanticFormControl(
        string elementName,
        string type,
        string name,
        IReadOnlyList<string> values,
        bool isChecked,
        bool isDisabled,
        bool isRequired,
        bool isReadOnly,
        bool isMultiple,
        string formOwnerId,
        string pattern,
        string minimum,
        string maximum,
        string step,
        int? minimumLength,
        int? maximumLength,
        string placeholder,
        string? formAction,
        string? formMethod,
        string? formEncodingType,
        string? formTarget,
        bool formNoValidate) {
        ElementName = elementName;
        Type = type;
        Name = name;
        Values = values;
        IsChecked = isChecked;
        IsDisabled = isDisabled;
        IsRequired = isRequired;
        IsReadOnly = isReadOnly;
        IsMultiple = isMultiple;
        FormOwnerId = formOwnerId;
        Pattern = pattern;
        Minimum = minimum;
        Maximum = maximum;
        Step = step;
        MinimumLength = minimumLength;
        MaximumLength = maximumLength;
        Placeholder = placeholder;
        FormAction = formAction;
        FormMethod = formMethod;
        FormEncodingType = formEncodingType;
        FormTarget = formTarget;
        FormNoValidate = formNoValidate;
    }

    /// <summary>Normalized source element name, such as input, select, textarea, or button.</summary>
    public string ElementName { get; }

    /// <summary>Effective HTML control type, including specification defaults and invalid-value fallback.</summary>
    public string Type { get; }

    /// <summary>Control name.</summary>
    public string Name { get; }

    /// <summary>Current values; multi-select controls can expose more than one value.</summary>
    public IReadOnlyList<string> Values { get; }

    /// <summary>First current value, or an empty string when the control has no value.</summary>
    public string Value => Values.Count == 0 ? string.Empty : Values[0];

    /// <summary>Whether a checkbox or radio control is checked.</summary>
    public bool IsChecked { get; }

    /// <summary>Whether the control is disabled directly or by an ancestor disabled fieldset.</summary>
    public bool IsDisabled { get; }

    /// <summary>Whether the control declares required validation.</summary>
    public bool IsRequired { get; }

    /// <summary>Whether the control declares readonly state.</summary>
    public bool IsReadOnly { get; }

    /// <summary>Whether the control accepts multiple values.</summary>
    public bool IsMultiple { get; }

    /// <summary>Resolved form-owner id, or an empty string when no named owner is resolved.</summary>
    public string FormOwnerId { get; }

    /// <summary>Authored validation pattern.</summary>
    public string Pattern { get; }
    /// <summary>Authored minimum value.</summary>
    public string Minimum { get; }
    /// <summary>Authored maximum value.</summary>
    public string Maximum { get; }
    /// <summary>Authored value step.</summary>
    public string Step { get; }
    /// <summary>Authored minimum text length.</summary>
    public int? MinimumLength { get; }
    /// <summary>Authored maximum text length.</summary>
    public int? MaximumLength { get; }
    /// <summary>Authored placeholder text.</summary>
    public string Placeholder { get; }
    /// <summary>Submitter-specific action override; null for non-submitters or when absent.</summary>
    public string? FormAction { get; }
    /// <summary>Effective submitter-specific method override; null for non-submitters or when absent.</summary>
    public string? FormMethod { get; }
    /// <summary>Effective submitter-specific encoding override; null for non-submitters or when absent.</summary>
    public string? FormEncodingType { get; }
    /// <summary>Submitter-specific target override; null for non-submitters or when absent.</summary>
    public string? FormTarget { get; }
    /// <summary>Whether this submitter disables form validation.</summary>
    public bool FormNoValidate { get; }
}
