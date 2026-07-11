namespace OfficeIMO.Pdf;

/// <summary>DocMDP permission level declared by a certification signature.</summary>
public enum PdfCertificationPermissionLevel {
    /// <summary>No changes are permitted after certification (`/P 1`).</summary>
    NoChanges = 1,

    /// <summary>Form filling and additional signatures are permitted (`/P 2`).</summary>
    FormFillingAndSignatures = 2,

    /// <summary>Form filling, annotations, and additional signatures are permitted (`/P 3`).</summary>
    FormFillingAnnotationsAndSignatures = 3
}
