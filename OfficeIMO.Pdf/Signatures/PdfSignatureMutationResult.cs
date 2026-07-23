namespace OfficeIMO.Pdf;

/// <summary>Before/after structural preservation state for one pre-existing PDF signature.</summary>
public sealed class PdfSignatureMutationResult {
    internal PdfSignatureMutationResult(
        PdfSignatureValidationResult before,
        PdfSignatureValidationResult? after,
        int? signedRevisionNumberBefore,
        int? signedRevisionNumberAfter,
        IReadOnlyList<int> coveredRevisionsBefore,
        IReadOnlyList<int> coveredRevisionsAfter,
        bool originalBytesPreserved,
        bool byteRangePreserved,
        bool activeDefinitionPreserved,
        bool hasLaterRevisionsBefore,
        bool hasLaterRevisionsAfter,
        PdfSignatureMutationPermissionStatus permissionStatus) {
        Before = before;
        After = after;
        SignedRevisionNumberBefore = signedRevisionNumberBefore;
        SignedRevisionNumberAfter = signedRevisionNumberAfter;
        CoveredRevisionsBefore = coveredRevisionsBefore;
        CoveredRevisionsAfter = coveredRevisionsAfter;
        OriginalBytesPreserved = originalBytesPreserved;
        ByteRangePreserved = byteRangePreserved;
        ActiveDefinitionPreserved = activeDefinitionPreserved;
        HasLaterRevisionsBefore = hasLaterRevisionsBefore;
        HasLaterRevisionsAfter = hasLaterRevisionsAfter;
        PermissionStatus = permissionStatus;
    }

    /// <summary>Structural validation result before the mutation.</summary>
    public PdfSignatureValidationResult Before { get; }

    /// <summary>Matching structural validation result after the mutation, when still present.</summary>
    public PdfSignatureValidationResult? After { get; }

    /// <summary>True when the signature object is still readable after the mutation.</summary>
    public bool IsPresentAfter => After is not null;

    /// <summary>Revision whose end is covered by the original byte range, when determinable.</summary>
    public int? SignedRevisionNumberBefore { get; }

    /// <summary>Revision whose end is covered by the post-mutation byte range, when determinable.</summary>
    public int? SignedRevisionNumberAfter { get; }

    /// <summary>Revision numbers fully covered before the mutation.</summary>
    public IReadOnlyList<int> CoveredRevisionsBefore { get; }

    /// <summary>Revision numbers fully covered after the mutation.</summary>
    public IReadOnlyList<int> CoveredRevisionsAfter { get; }

    /// <summary>True when the complete original file remains an exact prefix of the output.</summary>
    public bool OriginalBytesPreserved { get; }

    /// <summary>True when the signature's exact /ByteRange values are unchanged.</summary>
    public bool ByteRangePreserved { get; }

    /// <summary>True when the active signature and owning field object graphs are unchanged.</summary>
    public bool ActiveDefinitionPreserved { get; }

    /// <summary>True when revisions or unsigned bytes followed the signature before the mutation.</summary>
    public bool HasLaterRevisionsBefore { get; }

    /// <summary>True when revisions or unsigned bytes follow the signature after the mutation.</summary>
    public bool HasLaterRevisionsAfter { get; }

    /// <summary>DocMDP/FieldMDP-aware structural permission result for the requested mutation.</summary>
    public PdfSignatureMutationPermissionStatus PermissionStatus { get; }

    /// <summary>True when structural signature state remains present, byte-identical, and no less valid after mutation.</summary>
    public bool IsStructurallyPreserved =>
        IsPresentAfter &&
        OriginalBytesPreserved &&
        ByteRangePreserved &&
        ActiveDefinitionPreserved &&
        Before.IsStructurallyValid == After!.IsStructurallyValid;
}
