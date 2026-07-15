namespace OfficeIMO.Pdf;

/// <summary>Standard embedded-file property exposed as a document portfolio field.</summary>
public enum PdfPortfolioFieldKind {
    /// <summary>Embedded file name.</summary>
    FileName,
    /// <summary>Embedded file description.</summary>
    Description,
    /// <summary>Embedded file creation date.</summary>
    CreationDate,
    /// <summary>Embedded file modification date.</summary>
    ModificationDate,
    /// <summary>Embedded file size.</summary>
    Size
}
