namespace OfficeIMO.Pdf;

/// <summary>Resource budgets applied while parsing PDF syntax and object graphs.</summary>
public sealed class PdfReadLimits {
    internal const int DefaultMaxDecodedStreamBytes = 256 * 1024 * 1024;
    internal const int DefaultMaxContentOperations = 1_000_000;
    internal const int DefaultMaxContentOperands = 1_000_000;
    internal const int DefaultMaxContentNestingDepth = 128;

    /// <summary>Creates default parser budgets that callers can customize without changing another options instance.</summary>
    public static PdfReadLimits Default => new PdfReadLimits();

    /// <summary>Maximum input byte count accepted before text/object scanning. Default: 512 MiB.</summary>
    public long MaxInputBytes { get; init; } = 512L * 1024L * 1024L;

    /// <summary>Maximum number of indirect object declarations accepted. Default: 500,000.</summary>
    public int MaxIndirectObjects { get; init; } = 500_000;

    /// <summary>Maximum raw byte count allocated for one stream. Default: 256 MiB.</summary>
    public int MaxRawStreamBytes { get; init; } = 256 * 1024 * 1024;

    /// <summary>Maximum decoded byte count produced from one filtered stream. Default: 256 MiB.</summary>
    public int MaxDecodedStreamBytes { get; init; } = DefaultMaxDecodedStreamBytes;

    /// <summary>Maximum characters tokenized from one object or dictionary. Default: 1,000,000.</summary>
    public int MaxObjectCharacters { get; init; } = 1_000_000;

    /// <summary>Maximum syntax tokens accepted in one object or dictionary. Default: 100,000.</summary>
    public int MaxTokensPerObject { get; init; } = 100_000;

    /// <summary>Maximum nested array/dictionary depth accepted by the object parser. Default: 128.</summary>
    public int MaxObjectNestingDepth { get; init; } = 128;

    /// <summary>Maximum wall-clock time spent in the core object parsing pass. Default: 30 seconds.</summary>
    public TimeSpan MaxObjectParsingTime { get; init; } = TimeSpan.FromSeconds(30);

    /// <summary>Maximum cross-reference revisions discovered in one input. Default: 10,000.</summary>
    public int MaxRevisions { get; init; } = 10_000;

    /// <summary>Maximum page-tree dictionaries traversed. Default: 100,000.</summary>
    public int MaxPageTreeNodes { get; init; } = 100_000;

    /// <summary>Maximum nested page-tree depth. Default: 1,024.</summary>
    public int MaxPageTreeDepth { get; init; } = 1_024;

    /// <summary>Maximum pages discovered in one document. Default: 100,000.</summary>
    public int MaxPages { get; init; } = 100_000;

    /// <summary>Maximum AcroForm field-tree nodes or terminal fields. Default: 100,000.</summary>
    public int MaxFormFields { get; init; } = 100_000;

    /// <summary>Maximum nested AcroForm field-tree depth. Default: 256.</summary>
    public int MaxFormFieldDepth { get; init; } = 256;

    /// <summary>Maximum annotations declared on one page. Default: 100,000.</summary>
    public int MaxAnnotationsPerPage { get; init; } = 100_000;

    /// <summary>Maximum operators parsed from one page or form content stream. Default: 1,000,000.</summary>
    public int MaxContentOperations { get; init; } = DefaultMaxContentOperations;

    /// <summary>Maximum operand values and dictionary keys parsed from one page or form content stream. Default: 1,000,000.</summary>
    public int MaxContentOperands { get; init; } = DefaultMaxContentOperands;

    /// <summary>Maximum nested lexical arrays/dictionaries or form XObjects while parsing page content. Default: 128.</summary>
    public int MaxContentNestingDepth { get; init; } = DefaultMaxContentNestingDepth;

    internal PdfReadLimits WithMinimumInputBytes(long minimumInputBytes) {
        return new PdfReadLimits {
            MaxInputBytes = Math.Max(MaxInputBytes, minimumInputBytes),
            MaxIndirectObjects = MaxIndirectObjects,
            MaxRawStreamBytes = MaxRawStreamBytes,
            MaxDecodedStreamBytes = MaxDecodedStreamBytes,
            MaxObjectCharacters = MaxObjectCharacters,
            MaxTokensPerObject = MaxTokensPerObject,
            MaxObjectNestingDepth = MaxObjectNestingDepth,
            MaxObjectParsingTime = MaxObjectParsingTime,
            MaxRevisions = MaxRevisions,
            MaxPageTreeNodes = MaxPageTreeNodes,
            MaxPageTreeDepth = MaxPageTreeDepth,
            MaxPages = MaxPages,
            MaxFormFields = MaxFormFields,
            MaxFormFieldDepth = MaxFormFieldDepth,
            MaxAnnotationsPerPage = MaxAnnotationsPerPage,
            MaxContentOperations = MaxContentOperations,
            MaxContentOperands = MaxContentOperands,
            MaxContentNestingDepth = MaxContentNestingDepth
        };
    }

    internal void Validate() {
        if (MaxInputBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxInputBytes), MaxInputBytes, "Maximum input bytes must be positive.");
        }

        if (MaxIndirectObjects <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxIndirectObjects), MaxIndirectObjects, "Maximum indirect objects must be positive.");
        }

        if (MaxRawStreamBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxRawStreamBytes), MaxRawStreamBytes, "Maximum raw stream bytes must be positive.");
        }

        if (MaxDecodedStreamBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxDecodedStreamBytes), MaxDecodedStreamBytes, "Maximum decoded stream bytes must be positive.");
        }

        if (MaxObjectCharacters <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxObjectCharacters), MaxObjectCharacters, "Maximum object characters must be positive.");
        }

        if (MaxTokensPerObject <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxTokensPerObject), MaxTokensPerObject, "Maximum tokens per object must be positive.");
        }

        if (MaxObjectNestingDepth <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaxObjectNestingDepth), MaxObjectNestingDepth, "Maximum object nesting depth must be positive.");
        }

        if (MaxObjectParsingTime <= TimeSpan.Zero) {
            throw new ArgumentOutOfRangeException(nameof(MaxObjectParsingTime), MaxObjectParsingTime, "Maximum object parsing time must be positive.");
        }

        ValidatePositive(MaxRevisions, nameof(MaxRevisions), "Maximum revisions must be positive.");
        ValidatePositive(MaxPageTreeNodes, nameof(MaxPageTreeNodes), "Maximum page-tree nodes must be positive.");
        ValidatePositive(MaxPageTreeDepth, nameof(MaxPageTreeDepth), "Maximum page-tree depth must be positive.");
        ValidatePositive(MaxPages, nameof(MaxPages), "Maximum pages must be positive.");
        ValidatePositive(MaxFormFields, nameof(MaxFormFields), "Maximum form fields must be positive.");
        ValidatePositive(MaxFormFieldDepth, nameof(MaxFormFieldDepth), "Maximum form-field depth must be positive.");
        ValidatePositive(MaxAnnotationsPerPage, nameof(MaxAnnotationsPerPage), "Maximum annotations per page must be positive.");
        ValidatePositive(MaxContentOperations, nameof(MaxContentOperations), "Maximum content operations must be positive.");
        ValidatePositive(MaxContentOperands, nameof(MaxContentOperands), "Maximum content operands must be positive.");
        ValidatePositive(MaxContentNestingDepth, nameof(MaxContentNestingDepth), "Maximum content nesting depth must be positive.");
    }

    private static void ValidatePositive(int value, string parameterName, string message) {
        if (value <= 0) {
            throw new ArgumentOutOfRangeException(parameterName, value, message);
        }
    }
}
