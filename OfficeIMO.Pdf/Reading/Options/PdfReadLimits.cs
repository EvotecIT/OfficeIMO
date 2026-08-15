namespace OfficeIMO.Pdf;

/// <summary>Resource budgets applied while parsing PDF syntax and object graphs.</summary>
public sealed class PdfReadLimits {
    internal const int DefaultMaxDecodedStreamBytes = 256 * 1024 * 1024;
    internal const long DefaultMaxTotalDecodedStreamBytes = 512L * 1024L * 1024L;
    internal const int DefaultMaxContentOperations = 1_000_000;
    internal const int DefaultMaxContentOperands = 1_000_000;
    internal const int DefaultMaxContentNestingDepth = 128;
    internal const int DefaultMaxPageContentBytes = 256 * 1024 * 1024;
    internal const long DefaultMaxRetainedContentBytes = 512L * 1024L * 1024L;
    internal const int DefaultMaxActualTextCharacters = 1_000_000;
    internal const int DefaultMaxDecodedTextCharacters = 10_000_000;
    internal const int DefaultMaxTextSearchMatches = 100_000;
    internal const int DefaultMaxNameTreeNodes = 100_000;
    internal const int DefaultMaxNameTreeDepth = 128;
    internal const int DefaultMaxJavaScriptBytes = 4_000_000;
    internal const int DefaultMaxJavaScripts = 10_000;
    internal const int DefaultMaxWidgetActions = 100_000;
    internal const long DefaultMaxTotalJavaScriptBytes = 32L * 1024L * 1024L;
    internal const int DefaultMaxAttachments = 100_000;
    internal const long DefaultMaxTotalAttachmentBytes = 256L * 1024L * 1024L;
    internal const int DefaultMaxType3GlyphInvocationsPerPage = 1_000_000;

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

    /// <summary>Maximum aggregate decoded stream bytes cached while parsing one document. Default: 512 MiB.</summary>
    internal long MaxTotalDecodedStreamBytes { get; init; } = DefaultMaxTotalDecodedStreamBytes;

    /// <summary>Maximum aggregate decoded content-stream bytes materialized for one page. Default: 256 MiB.</summary>
    public int MaxPageContentBytes { get; init; } = DefaultMaxPageContentBytes;

    /// <summary>Maximum aggregate decoded content-stream bytes retained by one document-wide validation operation. Default: 512 MiB.</summary>
    public long MaxRetainedContentBytes { get; init; } = DefaultMaxRetainedContentBytes;

    /// <summary>Maximum characters emitted from marked-content ActualText replacements on one page, including nested Form XObjects. Default: 1,000,000.</summary>
    public int MaxActualTextCharacters { get; init; } = DefaultMaxActualTextCharacters;

    /// <summary>Maximum font-decoded text characters emitted on one page, including nested Form XObjects. Default: 10,000,000.</summary>
    public int MaxDecodedTextCharacters { get; init; } = DefaultMaxDecodedTextCharacters;

    /// <summary>Maximum text-search matches materialized by one Find or ReplaceAll operation. Default: 100,000.</summary>
    public int MaxTextSearchMatches { get; init; } = DefaultMaxTextSearchMatches;

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

    /// <summary>Maximum indirect nodes traversed in one PDF name tree. Default: 100,000.</summary>
    public int MaxNameTreeNodes { get; init; } = DefaultMaxNameTreeNodes;

    /// <summary>Maximum nested PDF name-tree depth. Default: 128.</summary>
    public int MaxNameTreeDepth { get; init; } = DefaultMaxNameTreeDepth;

    /// <summary>Maximum decoded source bytes retained for one named or widget JavaScript action. Default: 4,000,000.</summary>
    public int MaxJavaScriptBytes { get; init; } = DefaultMaxJavaScriptBytes;

    /// <summary>Maximum JavaScript entries discovered in one PDF action surface. Default: 10,000.</summary>
    public int MaxJavaScripts { get; init; } = DefaultMaxJavaScripts;

    /// <summary>Maximum widget action nodes materialized while reading AcroForm action graphs. Default: 100,000.</summary>
    public int MaxWidgetActions { get; init; } = DefaultMaxWidgetActions;

    /// <summary>Maximum aggregate decoded source bytes retained for one PDF JavaScript action surface. Default: 32 MiB.</summary>
    public long MaxTotalJavaScriptBytes { get; init; } = DefaultMaxTotalJavaScriptBytes;

    /// <summary>Maximum attachment records discovered across name trees, associated files, and annotations. Default: 100,000.</summary>
    public int MaxAttachments { get; init; } = DefaultMaxAttachments;

    /// <summary>Maximum aggregate decoded bytes retained for unique embedded attachment streams. Default: 256 MiB.</summary>
    public long MaxTotalAttachmentBytes { get; init; } = DefaultMaxTotalAttachmentBytes;

    /// <summary>Maximum named appearance states declared for one AcroForm widget. Default: 4,096.</summary>
    public int MaxFormFieldAppearanceStates { get; init; } = 4_096;

    /// <summary>Maximum annotations declared on one page. Default: 100,000.</summary>
    public int MaxAnnotationsPerPage { get; init; } = 100_000;

    /// <summary>Maximum named color-space resources inspected on one page. Default: 4,096.</summary>
    public int MaxColorSpaceResourcesPerPage { get; init; } = 4_096;

    /// <summary>Maximum operators parsed from one page or form content stream. Default: 1,000,000.</summary>
    public int MaxContentOperations { get; init; } = DefaultMaxContentOperations;

    /// <summary>Maximum operand values and dictionary keys parsed from one page or form content stream. Default: 1,000,000.</summary>
    public int MaxContentOperands { get; init; } = DefaultMaxContentOperands;

    /// <summary>Maximum nested lexical arrays/dictionaries or form XObjects while parsing page content. Default: 128.</summary>
    public int MaxContentNestingDepth { get; init; } = DefaultMaxContentNestingDepth;

    /// <summary>Maximum Type 3 glyph programs invoked while rendering one page, including nested forms. Default: 1,000,000.</summary>
    public int MaxType3GlyphInvocationsPerPage { get; init; } = DefaultMaxType3GlyphInvocationsPerPage;

    internal PdfReadLimits WithMinimumInputBytes(long minimumInputBytes) {
        return new PdfReadLimits {
            MaxInputBytes = Math.Max(MaxInputBytes, minimumInputBytes),
            MaxIndirectObjects = MaxIndirectObjects,
            MaxRawStreamBytes = MaxRawStreamBytes,
            MaxDecodedStreamBytes = MaxDecodedStreamBytes,
            MaxTotalDecodedStreamBytes = MaxTotalDecodedStreamBytes,
            MaxPageContentBytes = MaxPageContentBytes,
            MaxRetainedContentBytes = MaxRetainedContentBytes,
            MaxActualTextCharacters = MaxActualTextCharacters,
            MaxDecodedTextCharacters = MaxDecodedTextCharacters,
            MaxTextSearchMatches = MaxTextSearchMatches,
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
            MaxNameTreeNodes = MaxNameTreeNodes,
            MaxNameTreeDepth = MaxNameTreeDepth,
            MaxJavaScriptBytes = MaxJavaScriptBytes,
            MaxJavaScripts = MaxJavaScripts,
            MaxWidgetActions = MaxWidgetActions,
            MaxTotalJavaScriptBytes = MaxTotalJavaScriptBytes,
            MaxAttachments = MaxAttachments,
            MaxTotalAttachmentBytes = MaxTotalAttachmentBytes,
            MaxFormFieldAppearanceStates = MaxFormFieldAppearanceStates,
            MaxAnnotationsPerPage = MaxAnnotationsPerPage,
            MaxColorSpaceResourcesPerPage = MaxColorSpaceResourcesPerPage,
            MaxContentOperations = MaxContentOperations,
            MaxContentOperands = MaxContentOperands,
            MaxContentNestingDepth = MaxContentNestingDepth,
            MaxType3GlyphInvocationsPerPage = MaxType3GlyphInvocationsPerPage
        };
    }

    internal PdfReadLimits WithMaximumContainerEntries(
        int maximumContainerEntries,
        long? maximumDecodedStreamBytes = null,
        long? maximumTotalDecodedStreamBytes = null,
        long? maximumTotalAttachmentBytes = null,
        bool preserveExistingDecodedStreamLimit = true) {
        if (maximumContainerEntries <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maximumContainerEntries), maximumContainerEntries, "Maximum container entries must be positive.");
        }
        if (maximumDecodedStreamBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maximumDecodedStreamBytes), maximumDecodedStreamBytes, "Maximum decoded stream bytes must be positive.");
        }
        if (maximumTotalDecodedStreamBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maximumTotalDecodedStreamBytes), maximumTotalDecodedStreamBytes, "Maximum aggregate decoded stream bytes must be positive.");
        }
        if (maximumTotalAttachmentBytes <= 0) {
            throw new ArgumentOutOfRangeException(nameof(maximumTotalAttachmentBytes), maximumTotalAttachmentBytes, "Maximum aggregate attachment bytes must be positive.");
        }
        int requestedDecodedStreamBytes = maximumDecodedStreamBytes.HasValue
            ? (int)Math.Min(maximumDecodedStreamBytes.Value, int.MaxValue)
            : MaxDecodedStreamBytes;
        int effectiveDecodedStreamBytes = preserveExistingDecodedStreamLimit
            ? Math.Min(MaxDecodedStreamBytes, requestedDecodedStreamBytes)
            : requestedDecodedStreamBytes;
        return new PdfReadLimits {
            MaxInputBytes = MaxInputBytes,
            MaxIndirectObjects = Math.Min(MaxIndirectObjects, maximumContainerEntries),
            MaxRawStreamBytes = MaxRawStreamBytes,
            MaxDecodedStreamBytes = effectiveDecodedStreamBytes,
            MaxTotalDecodedStreamBytes = maximumTotalDecodedStreamBytes ?? maximumDecodedStreamBytes
                ?? MaxTotalDecodedStreamBytes,
            MaxPageContentBytes = MaxPageContentBytes,
            MaxRetainedContentBytes = MaxRetainedContentBytes,
            MaxActualTextCharacters = MaxActualTextCharacters,
            MaxDecodedTextCharacters = MaxDecodedTextCharacters,
            MaxTextSearchMatches = MaxTextSearchMatches,
            MaxObjectCharacters = MaxObjectCharacters,
            MaxTokensPerObject = MaxTokensPerObject,
            MaxObjectNestingDepth = MaxObjectNestingDepth,
            MaxObjectParsingTime = MaxObjectParsingTime,
            MaxRevisions = Math.Min(MaxRevisions, maximumContainerEntries),
            MaxPageTreeNodes = Math.Min(MaxPageTreeNodes, maximumContainerEntries),
            MaxPageTreeDepth = MaxPageTreeDepth,
            MaxPages = Math.Min(MaxPages, maximumContainerEntries),
            MaxFormFields = Math.Min(MaxFormFields, maximumContainerEntries),
            MaxFormFieldDepth = MaxFormFieldDepth,
            MaxNameTreeNodes = Math.Min(MaxNameTreeNodes, maximumContainerEntries),
            MaxNameTreeDepth = MaxNameTreeDepth,
            MaxJavaScriptBytes = MaxJavaScriptBytes,
            MaxJavaScripts = Math.Min(MaxJavaScripts, maximumContainerEntries),
            MaxWidgetActions = Math.Min(MaxWidgetActions, maximumContainerEntries),
            MaxTotalJavaScriptBytes = MaxTotalJavaScriptBytes,
            MaxAttachments = Math.Min(MaxAttachments, maximumContainerEntries),
            MaxTotalAttachmentBytes = maximumTotalAttachmentBytes ?? MaxTotalAttachmentBytes,
            MaxFormFieldAppearanceStates = Math.Min(MaxFormFieldAppearanceStates, maximumContainerEntries),
            MaxAnnotationsPerPage = Math.Min(MaxAnnotationsPerPage, maximumContainerEntries),
            MaxColorSpaceResourcesPerPage = Math.Min(MaxColorSpaceResourcesPerPage, maximumContainerEntries),
            MaxContentOperations = MaxContentOperations,
            MaxContentOperands = MaxContentOperands,
            MaxContentNestingDepth = MaxContentNestingDepth,
            MaxType3GlyphInvocationsPerPage = MaxType3GlyphInvocationsPerPage
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

        ValidatePositive(MaxPageContentBytes, nameof(MaxPageContentBytes), "Maximum aggregate page content bytes must be positive.");
        if (MaxRetainedContentBytes <= 0L) {
            throw new ArgumentOutOfRangeException(nameof(MaxRetainedContentBytes), MaxRetainedContentBytes, "Maximum retained content bytes must be positive.");
        }
        ValidatePositive(MaxActualTextCharacters, nameof(MaxActualTextCharacters), "Maximum ActualText characters must be positive.");
        ValidatePositive(MaxDecodedTextCharacters, nameof(MaxDecodedTextCharacters), "Maximum decoded text characters must be positive.");
        ValidatePositive(MaxTextSearchMatches, nameof(MaxTextSearchMatches), "Maximum text-search matches must be positive.");

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
        ValidatePositive(MaxNameTreeNodes, nameof(MaxNameTreeNodes), "Maximum name-tree nodes must be positive.");
        ValidatePositive(MaxNameTreeDepth, nameof(MaxNameTreeDepth), "Maximum name-tree depth must be positive.");
        ValidatePositive(MaxJavaScriptBytes, nameof(MaxJavaScriptBytes), "Maximum document JavaScript bytes must be positive.");
        ValidatePositive(MaxJavaScripts, nameof(MaxJavaScripts), "Maximum document JavaScript entries must be positive.");
        ValidatePositive(MaxWidgetActions, nameof(MaxWidgetActions), "Maximum widget action nodes must be positive.");
        if (MaxTotalJavaScriptBytes <= 0L) {
            throw new ArgumentOutOfRangeException(nameof(MaxTotalJavaScriptBytes), MaxTotalJavaScriptBytes, "Maximum aggregate document JavaScript bytes must be positive.");
        }
        ValidatePositive(MaxAttachments, nameof(MaxAttachments), "Maximum attachments must be positive.");
        if (MaxTotalAttachmentBytes <= 0L) {
            throw new ArgumentOutOfRangeException(nameof(MaxTotalAttachmentBytes), MaxTotalAttachmentBytes, "Maximum aggregate attachment bytes must be positive.");
        }
        ValidatePositive(MaxFormFieldAppearanceStates, nameof(MaxFormFieldAppearanceStates), "Maximum form-field appearance states must be positive.");
        ValidatePositive(MaxAnnotationsPerPage, nameof(MaxAnnotationsPerPage), "Maximum annotations per page must be positive.");
        ValidatePositive(MaxColorSpaceResourcesPerPage, nameof(MaxColorSpaceResourcesPerPage), "Maximum color-space resources per page must be positive.");
        ValidatePositive(MaxContentOperations, nameof(MaxContentOperations), "Maximum content operations must be positive.");
        ValidatePositive(MaxContentOperands, nameof(MaxContentOperands), "Maximum content operands must be positive.");
        ValidatePositive(MaxContentNestingDepth, nameof(MaxContentNestingDepth), "Maximum content nesting depth must be positive.");
        ValidatePositive(MaxType3GlyphInvocationsPerPage, nameof(MaxType3GlyphInvocationsPerPage), "Maximum Type 3 glyph invocations per page must be positive.");
    }

    private static void ValidatePositive(int value, string parameterName, string message) {
        if (value <= 0) {
            throw new ArgumentOutOfRangeException(parameterName, value, message);
        }
    }
}
