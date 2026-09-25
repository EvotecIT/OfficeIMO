namespace OfficeIMO.Pdf;

/// <summary>Source features omitted when pages become content on new sheets.</summary>
[Flags]
public enum PdfImpositionSourceFeatureLoss {
    /// <summary>No known source feature was omitted.</summary>
    None = 0,
    /// <summary>Page annotations and their appearances were not transferred.</summary>
    Annotations = 1,
    /// <summary>AcroForm fields and widgets were not transferred.</summary>
    Forms = 2,
    /// <summary>Document structure tags were not transferred.</summary>
    StructureTags = 4,
    /// <summary>Source signatures were removed in an explicit unsigned derivative.</summary>
    Signatures = 8
}
