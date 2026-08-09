namespace OfficeIMO.OpenDocument;

/// <summary>Severity shown when an ODF spreadsheet validation fails.</summary>
public enum OdsValidationMessageType {
    /// <summary>Reject the value.</summary>
    Stop,
    /// <summary>Warn before accepting the value.</summary>
    Warning,
    /// <summary>Show an informational message.</summary>
    Information
}

/// <summary>An XML-backed workbook named range.</summary>
public sealed class OdsNamedRange {
    private readonly OdsDocument _document;
    private readonly XElement _element;

    internal OdsNamedRange(OdsDocument document, XElement element) { _document = document; _element = element; }

    /// <summary>Range name.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Table + "name") ?? string.Empty;
    /// <summary>ODF cell range address.</summary>
    public string CellRangeAddress {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "cell-range-address") ?? string.Empty;
        set { _element.SetAttributeValue(OdfNamespaces.Table + "cell-range-address", value); Dirty(); }
    }
    /// <summary>ODF base cell address.</summary>
    public string BaseCellAddress {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "base-cell-address") ?? string.Empty;
        set { _element.SetAttributeValue(OdfNamespaces.Table + "base-cell-address", value); Dirty(); }
    }

    private void Dirty() => _document.MarkPartDirty("content.xml");
}

/// <summary>An XML-backed spreadsheet content validation rule.</summary>
public sealed class OdsValidation {
    private readonly OdsDocument _document;
    private readonly XElement _element;

    internal OdsValidation(OdsDocument document, XElement element) { _document = document; _element = element; }

    /// <summary>Validation name.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Table + "name") ?? string.Empty;
    /// <summary>Preserved ODF validation condition.</summary>
    public string Condition {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "condition") ?? string.Empty;
        set { _element.SetAttributeValue(OdfNamespaces.Table + "condition", value); Dirty(); }
    }
    /// <summary>Typed interoperable condition, or null when the preserved expression is implementation-specific.</summary>
    public OdsValidationConditionSyntax? ParsedCondition {
        get => OdsValidationConditionSyntax.TryParse(Condition, out OdsValidationConditionSyntax? condition) ? condition : null;
        set => Condition = value?.ToString() ?? throw new ArgumentNullException(nameof(value));
    }
    /// <summary>Whether empty cells satisfy the rule.</summary>
    public bool AllowEmptyCell {
        get => (string?)_element.Attribute(OdfNamespaces.Table + "allow-empty-cell") != "false";
        set { _element.SetAttributeValue(OdfNamespaces.Table + "allow-empty-cell", value ? "true" : "false"); Dirty(); }
    }

    /// <summary>Sets the optional input help shown when a validated cell is selected.</summary>
    public void SetHelpMessage(string? title, string? text, bool display = true) {
        SetMessage(OdfNamespaces.Table + "help-message", title, text, display, null);
    }

    /// <summary>Sets the optional error shown when the entered value fails validation.</summary>
    public void SetErrorMessage(
        string? title,
        string? text,
        OdsValidationMessageType messageType = OdsValidationMessageType.Stop,
        bool display = true) {
        SetMessage(OdfNamespaces.Table + "error-message", title, text, display, FormatMessageType(messageType));
    }

    /// <summary>Input-help title, if present.</summary>
    public string? HelpTitle => (string?)_element.Element(OdfNamespaces.Table + "help-message")?
        .Attribute(OdfNamespaces.Table + "title");
    /// <summary>Whether an input-help element is present.</summary>
    public bool HasHelpMessage => _element.Element(OdfNamespaces.Table + "help-message") != null;
    /// <summary>Input-help text, if present.</summary>
    public string? HelpText => ReadMessageText(OdfNamespaces.Table + "help-message");
    /// <summary>Whether input help is displayed.</summary>
    public bool ShowHelpMessage => ReadDisplay(OdfNamespaces.Table + "help-message");
    /// <summary>Validation-error title, if present.</summary>
    public string? ErrorTitle => (string?)_element.Element(OdfNamespaces.Table + "error-message")?
        .Attribute(OdfNamespaces.Table + "title");
    /// <summary>Whether a validation-error element is present.</summary>
    public bool HasErrorMessage => _element.Element(OdfNamespaces.Table + "error-message") != null;
    /// <summary>Validation-error text, if present.</summary>
    public string? ErrorText => ReadMessageText(OdfNamespaces.Table + "error-message");
    /// <summary>Whether validation errors are displayed.</summary>
    public bool ShowErrorMessage => ReadDisplay(OdfNamespaces.Table + "error-message");
    /// <summary>Validation-error severity.</summary>
    public OdsValidationMessageType ErrorMessageType {
        get {
            string? value = (string?)_element.Element(OdfNamespaces.Table + "error-message")?
                .Attribute(OdfNamespaces.Table + "message-type");
            if (string.Equals(value, "warning", StringComparison.OrdinalIgnoreCase)) return OdsValidationMessageType.Warning;
            if (string.Equals(value, "information", StringComparison.OrdinalIgnoreCase)) return OdsValidationMessageType.Information;
            return OdsValidationMessageType.Stop;
        }
    }

    private void SetMessage(XName name, string? title, string? text, bool display, string? messageType) {
        XElement? message = _element.Element(name);
        if (title == null && text == null) {
            message?.Remove();
            Dirty();
            return;
        }
        if (message == null) {
            message = new XElement(name);
            _element.Add(message);
        }
        message.SetAttributeValue(OdfNamespaces.Table + "display", display ? "true" : "false");
        message.SetAttributeValue(OdfNamespaces.Table + "title", title);
        message.SetAttributeValue(OdfNamespaces.Table + "message-type", messageType);
        message.RemoveNodes();
        if (text != null) message.Add(new XElement(OdfNamespaces.Text + "p", text));
        Dirty();
    }

    private string? ReadMessageText(XName name) {
        XElement? message = _element.Element(name);
        if (message == null) return null;
        return string.Join("\n", message.Elements(OdfNamespaces.Text + "p").Select(paragraph => paragraph.Value));
    }

    private bool ReadDisplay(XName name) =>
        string.Equals((string?)_element.Element(name)?.Attribute(OdfNamespaces.Table + "display"), "true", StringComparison.OrdinalIgnoreCase);

    private static string FormatMessageType(OdsValidationMessageType value) => value switch {
        OdsValidationMessageType.Stop => "stop",
        OdsValidationMessageType.Warning => "warning",
        OdsValidationMessageType.Information => "information",
        _ => throw new ArgumentOutOfRangeException(nameof(value))
    };

    private void Dirty() => _document.MarkPartDirty("content.xml");
}
