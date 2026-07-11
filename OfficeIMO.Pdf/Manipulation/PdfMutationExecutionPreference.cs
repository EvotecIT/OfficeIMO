namespace OfficeIMO.Pdf;

/// <summary>Caller constraint applied when choosing a PDF mutation execution mode.</summary>
public enum PdfMutationExecutionPreference {
    /// <summary>Choose the safest available implemented path, preferring a permitted full rewrite for ordinary inputs.</summary>
    Automatic,

    /// <summary>Require a complete rewrite and block when the input or operation cannot use one safely.</summary>
    RequireFullRewrite,

    /// <summary>Require an append-only revision and block when the input or operation cannot use one safely.</summary>
    RequireAppendOnly
}
