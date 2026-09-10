namespace OfficeIMO.Invoicing;

/// <summary>Invoice note and optional UNTDID 4451 subject code.</summary>
public sealed class InvoiceNote {
    /// <summary>Creates a note.</summary>
    public InvoiceNote(string text, string? subjectCode = null) { Text = text; SubjectCode = subjectCode; }
    /// <summary>Note text.</summary>
    public string Text { get; set; }
    /// <summary>Subject code.</summary>
    public string? SubjectCode { get; set; }
}

/// <summary>Referenced invoice number and optional issue date.</summary>
public sealed class InvoiceReference {
    /// <summary>Creates a reference.</summary>
    public InvoiceReference(string number, DateTime? issueDate = null) { Number = number; IssueDate = issueDate; }
    /// <summary>Document number.</summary>
    public string Number { get; set; }
    /// <summary>Issue date.</summary>
    public DateTime? IssueDate { get; set; }
}

/// <summary>Service or invoice period.</summary>
public sealed class InvoicePeriod {
    /// <summary>Inclusive start date.</summary>
    public DateTime? Start { get; set; }
    /// <summary>Inclusive end date.</summary>
    public DateTime? End { get; set; }
}

/// <summary>Supporting document, represented by a reference, external URI or embedded bytes.</summary>
public sealed class InvoiceSupportingDocument {
    /// <summary>Document reference (BT-122).</summary>
    public string Reference { get; set; } = string.Empty;
    /// <summary>Document description (BT-123).</summary>
    public string? Description { get; set; }
    /// <summary>External document location (BT-124); the engine never fetches it.</summary>
    public string? ExternalUri { get; set; }
    /// <summary>Embedded document bytes (BT-125).</summary>
    public byte[]? Data { get; set; }
    /// <summary>Embedded file name.</summary>
    public string? FileName { get; set; }
    /// <summary>Embedded file media type.</summary>
    public string? MimeType { get; set; }
}

/// <summary>Delivery information separate from the buyer's legal address.</summary>
public sealed class InvoiceDelivery {
    /// <summary>Recipient's name (BT-70).</summary>
    public string? Name { get; set; }
    /// <summary>Delivery location identifier (BT-71).</summary>
    public InvoiceIdentifier? LocationIdentifier { get; set; }
    /// <summary>Actual delivery date (BT-72).</summary>
    public DateTime? Date { get; set; }
    /// <summary>Delivery postal address (BG-15).</summary>
    public InvoiceAddress? Address { get; set; }
}
