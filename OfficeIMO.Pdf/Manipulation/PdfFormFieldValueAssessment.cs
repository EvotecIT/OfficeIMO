namespace OfficeIMO.Pdf;

/// <summary>Metadata-based feedback for a proposed interactive form value.</summary>
/// <remarks>This assessment does not execute JavaScript or replace document permissions, mutation planning,
/// appearance generation, or saved-output validation. Required fields produce warnings so unfinished forms can be saved.</remarks>
public sealed class PdfFormFieldValueAssessment {
    private PdfFormFieldValueAssessment(IReadOnlyList<PdfFormFieldValueIssue> issues) { Issues = issues; }

    /// <summary>Issues found in the field's declared metadata and proposed value.</summary>
    public IReadOnlyList<PdfFormFieldValueIssue> Issues { get; }
    /// <summary>Whether the proposed interactive edit conflicts with declared field constraints.</summary>
    public bool HasErrors => Issues.Any(issue => issue.IsError);

    /// <summary>Assesses a value without changing the field or document. Text limits count Unicode scalar values.</summary>
    public static PdfFormFieldValueAssessment Assess(PdfFormField field, PdfFormFieldValue value) {
        Guard.NotNull(field, nameof(field));
        Guard.NotNull(value, nameof(value));
        var issues = new List<PdfFormFieldValueIssue>();
        string[] values = value.Values.ToArray();
        if (field.IsReadOnly) Add(PdfFormFieldValueIssueCode.ReadOnly, true, "This field is read-only.");
        if (field.IsSignatureField || field.IsPushButton || field.Kind == PdfFormFieldKind.Unknown)
            Add(PdfFormFieldValueIssueCode.UnsupportedField, true, "This field does not accept an interactive form value.");
        if (values.Length > 1 && !(field.IsChoiceField && field.AllowsMultipleSelection))
            Add(PdfFormFieldValueIssueCode.MultipleValues, true, "This field accepts one value.");
        if (field.IsTextField && field.MaxLength is int maximum &&
            values.Any(text => PdfUnicodeScalarAnalysis.CountScalars(text) > maximum))
            Add(PdfFormFieldValueIssueCode.MaximumLength, true, $"This field allows at most {maximum} characters.");
        bool empty = values.All(text => string.IsNullOrEmpty(text) ||
            (field.IsCheckBox || field.IsRadioButton) && string.Equals(text, "Off", StringComparison.Ordinal));
        if (field.IsRequired && empty)
            Add(PdfFormFieldValueIssueCode.RequiredValue, false, "This required field is empty. You can save the form and complete it later.");
        if (field.IsRadioButton && field.IsNoToggleToOff && empty &&
            !string.IsNullOrEmpty(field.Value) && !string.Equals(field.Value, "Off", StringComparison.Ordinal))
            Add(PdfFormFieldValueIssueCode.NoToggleToOff, true, "Choose another radio option; this group does not allow clearing its selection.");
        if (field.IsChoiceField && !field.IsEditableChoice && values.Any(text => text.Length != 0 &&
            !field.Options.Any(option => option.ExportValue == text || option.DisplayText == text)))
            Add(PdfFormFieldValueIssueCode.UnknownChoice, true, "Choose a value offered by this field.");
        return new PdfFormFieldValueAssessment(issues.AsReadOnly());

        void Add(PdfFormFieldValueIssueCode code, bool error, string message) => issues.Add(new(code, error, message));
    }
}

/// <summary>A declared form-field constraint or incomplete required value.</summary>
public sealed class PdfFormFieldValueIssue {
    internal PdfFormFieldValueIssue(PdfFormFieldValueIssueCode code, bool isError, string message) {
        Code = code; IsError = isError; Message = message;
    }
    /// <summary>Stable reason suitable for host localization.</summary>
    public PdfFormFieldValueIssueCode Code { get; }
    /// <summary>True for a conflicting interactive edit; false for an incomplete required field.</summary>
    public bool IsError { get; }
    /// <summary>Human-readable fallback explanation.</summary>
    public string Message { get; }
}

/// <summary>Reasons reported by interactive form-value assessment.</summary>
public enum PdfFormFieldValueIssueCode {
    /// <summary>The field is read-only.</summary>
    ReadOnly,
    /// <summary>The field cannot be filled interactively.</summary>
    UnsupportedField,
    /// <summary>A scalar field received multiple values.</summary>
    MultipleValues,
    /// <summary>The proposed text exceeds the declared maximum length.</summary>
    MaximumLength,
    /// <summary>A required field has no value.</summary>
    RequiredValue,
    /// <summary>A selected radio group cannot be cleared.</summary>
    NoToggleToOff,
    /// <summary>The proposed choice is not offered by a non-editable field.</summary>
    UnknownChoice
}
