namespace OfficeIMO.Pdf;

/// <summary>Outcome of one independently reportable signature-validation dimension.</summary>
public enum PdfCryptographicValidationStatus {
    /// <summary>The provider did not perform this validation dimension.</summary>
    NotPerformed = 0,

    /// <summary>The validation dimension succeeded.</summary>
    Valid = 1,

    /// <summary>The validation dimension completed and found invalid evidence.</summary>
    Invalid = 2,

    /// <summary>Available evidence was insufficient for a definitive result.</summary>
    Indeterminate = 3,

    /// <summary>The validation dimension could not complete because of an error.</summary>
    Error = 4
}
