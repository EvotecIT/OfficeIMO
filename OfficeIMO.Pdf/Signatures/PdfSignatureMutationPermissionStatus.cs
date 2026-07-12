namespace OfficeIMO.Pdf;

/// <summary>Structural permission result for a requested mutation relative to an existing signature.</summary>
public enum PdfSignatureMutationPermissionStatus {
    /// <summary>The input contains no signature to constrain the mutation.</summary>
    NotApplicableUnsigned,

    /// <summary>The mutation planner and DocMDP/FieldMDP checks permit the requested append-only change.</summary>
    Permitted,

    /// <summary>The mutation planner or DocMDP/FieldMDP checks forbid the requested change.</summary>
    Forbidden,

    /// <summary>The available signature or revision evidence is insufficient to classify permission.</summary>
    Indeterminate
}
