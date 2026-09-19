namespace OfficeIMO.Invoicing.Pdf;

/// <summary>
/// Describes a visible approval block. This is presentation metadata and is not a PDF digital signature.
/// </summary>
public sealed class InvoicePdfApproval {
    /// <summary>Creates a visible approval block.</summary>
    public InvoicePdfApproval(string label, string name, string? role = null, DateTime? date = null) {
        Label = string.IsNullOrWhiteSpace(label) ? throw new ArgumentException("An approval label is required.", nameof(label)) : label;
        Name = string.IsNullOrWhiteSpace(name) ? throw new ArgumentException("An approver name is required.", nameof(name)) : name;
        Role = role;
        Date = date;
    }

    /// <summary>Caption such as <c>Prepared by</c> or <c>Approved by</c>.</summary>
    public string Label { get; }

    /// <summary>Displayed person name.</summary>
    public string Name { get; }

    /// <summary>Optional role or team.</summary>
    public string? Role { get; }

    /// <summary>Optional approval date.</summary>
    public DateTime? Date { get; }

    internal InvoicePdfApproval Snapshot() => new InvoicePdfApproval(Label, Name, Role, Date);
}
