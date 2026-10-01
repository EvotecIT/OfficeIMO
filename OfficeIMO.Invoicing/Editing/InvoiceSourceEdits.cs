namespace OfficeIMO.Invoicing;

/// <summary>Captures up to five independent source-field replacements. Null means retain the field; no field is inserted or removed.</summary>
public sealed class InvoiceSourceEdits {
    /// <summary>Creates a captured edit request. Dates use their calendar day; time and timezone are ignored. Text replacements must contain 1 to 4096 characters.</summary>
    public InvoiceSourceEdits(string? number = null, DateTime? issueDate = null, DateTime? dueDate = null, string? buyerReference = null, string? paymentReference = null) {
        Number = Capture(number, nameof(number)); IssueDate = issueDate?.Date; DueDate = dueDate?.Date;
        BuyerReference = Capture(buyerReference, nameof(buyerReference)); PaymentReference = Capture(paymentReference, nameof(paymentReference));
    }
    /// <summary>Maximum characters in one replacement text value.</summary>
    public const int MaximumTextCharacters = 4096;
    /// <summary>Replacement invoice number. Payment/reference fields are changed only when separately supplied.</summary>
    public string? Number { get; }
    /// <summary>Replacement issue date.</summary>
    public DateTime? IssueDate { get; }
    /// <summary>Replacement payment due date.</summary>
    public DateTime? DueDate { get; }
    /// <summary>Replacement buyer reference.</summary>
    public string? BuyerReference { get; }
    /// <summary>Replacement remittance reference.</summary>
    public string? PaymentReference { get; }
    private static string? Capture(string? value, string name) {
        if (value == null) return null;
        if (string.IsNullOrWhiteSpace(value) || value.Length > MaximumTextCharacters) throw new ArgumentException("Replacement text requires 1 to 4096 nonempty characters.", name);
        System.Xml.XmlConvert.VerifyXmlChars(value);
        return value;
    }
}

/// <summary>Result of an atomic source edit. Failure retains the original and returns no edited document.</summary>
public sealed class InvoiceSourceEditResult {
    internal InvoiceSourceEditResult(InvoiceSourceDocument? document, IReadOnlyList<InvoiceDiagnostic> diagnostics) { Document = document; Diagnostics = Array.AsReadOnly(diagnostics.ToArray()); }
    /// <summary>Edited immutable source, or null when any requested field could not be replaced safely.</summary>
    public InvoiceSourceDocument? Document { get; }
    /// <summary>True when every requested replacement completed.</summary>
    public bool Succeeded => Document != null;
    /// <summary>Operation findings. They do not establish model or standards validity.</summary>
    public IReadOnlyList<InvoiceDiagnostic> Diagnostics { get; }
}

/// <summary>Applies bounded preservation-aware source edits without rebuilding a partial semantic model.</summary>
public static class InvoiceSourceEditor {
    /// <summary>Replaces only existing, unique plaintext fields. Rejects XML signatures, nested/annotated fields and unsupported dates before mutation. Number changes do not infer payment-reference changes.</summary>
    public static InvoiceSourceEditResult Apply(InvoiceSourceDocument source, InvoiceSourceEdits edits) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (edits == null) throw new ArgumentNullException(nameof(edits));
        var replacements = new List<InvoiceSourceScalarEdit>();
        void Add(InvoiceSourceField field, string? value) { if (value != null) replacements.Add(new InvoiceSourceScalarEdit(field, value)); }
        Add(InvoiceSourceField.DocumentId, edits.Number);
        Add(InvoiceSourceField.IssueDate, edits.IssueDate.HasValue ? source.FormatDate(edits.IssueDate.Value) : null);
        Add(InvoiceSourceField.DueDate, edits.DueDate.HasValue ? source.FormatDate(edits.DueDate.Value) : null);
        Add(InvoiceSourceField.BuyerReference, edits.BuyerReference);
        Add(InvoiceSourceField.PaymentReference, edits.PaymentReference);
        try {
            InvoiceSourceDocument output = replacements.Count == 0 ? source : source.Apply(replacements);
            return new InvoiceSourceEditResult(output, new[] { new InvoiceDiagnostic("INV-SOURCE-EDIT-VALIDATION-REQUIRED",
                "Source XML content is retained. Validate the edited business data and exact XML against the intended standards release; related fields are changed only when explicitly requested.", "Source", InvoiceDiagnosticSeverity.Information) });
        } catch (Exception exception) when (exception is InvalidDataException || exception is InvalidOperationException || exception is NotSupportedException || exception is System.Xml.XmlException || exception is ArgumentException) {
            string location = exception is InvoiceSourceFieldException fieldException ? fieldException.Field switch {
                InvoiceSourceField.DocumentId => "Number", InvoiceSourceField.IssueDate => "IssueDate", InvoiceSourceField.DueDate => "DueDate",
                InvoiceSourceField.BuyerReference => "BuyerReference", InvoiceSourceField.PaymentReference => "PaymentReference", _ => "Source"
            } : source.HasXmlSignature ? "Source.Signature" : "Source";
            return new InvoiceSourceEditResult(null, new[] { new InvoiceDiagnostic("INV-SOURCE-EDIT-BLOCKED", exception.Message, location) });
        }
    }
}
