namespace OfficeIMO.Pdf;

/// <summary>Stable stages recorded by <see cref="PdfPipelineReport"/>.</summary>
public enum PdfPipelineStepKind {
    /// <summary>A new authored PDF document was created.</summary>
    Create,

    /// <summary>An existing PDF artifact was opened.</summary>
    Open,

    /// <summary>An existing PDF artifact was transformed.</summary>
    Mutation,

    /// <summary>A final PDF artifact was generated or saved.</summary>
    Output
}
