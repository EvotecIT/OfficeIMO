namespace OfficeIMO.Pdf;

/// <summary>How OfficeIMO.Pdf should apply a requested mutation to an existing PDF.</summary>
public enum PdfMutationExecutionMode {
    /// <summary>Create a new complete PDF revision by rewriting supported document objects.</summary>
    FullRewrite,

    /// <summary>Preserve every existing input byte and append a new incremental revision.</summary>
    AppendOnly,

    /// <summary>Do not attempt the mutation because no currently proven execution path is available.</summary>
    Blocked
}
