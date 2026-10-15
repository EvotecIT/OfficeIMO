namespace OfficeIMO.Pdf;

public sealed partial class PdfReadLimits {
    /// <summary>Creates an independent snapshot of every parser and content resource budget.</summary>
    public PdfReadLimits Clone() => CloneWithMaximumInputBytes(null);

    internal PdfReadLimits CloneWithMaximumInputBytes(long? maximumInputBytes) {
        var snapshot = new PdfReadLimits {
            MaxInputBytes = maximumInputBytes ?? MaxInputBytes,
            MaxIndirectObjects = MaxIndirectObjects,
            MaxRawStreamBytes = MaxRawStreamBytes,
            MaxDecodedStreamBytes = MaxDecodedStreamBytes,
            MaxTotalDecodedStreamBytes = MaxTotalDecodedStreamBytes,
            MaxPageContentBytes = MaxPageContentBytes,
            MaxRetainedContentBytes = MaxRetainedContentBytes,
            MaxActualTextCharacters = MaxActualTextCharacters,
            MaxDecodedTextCharacters = MaxDecodedTextCharacters,
            MaxTextSearchMatches = MaxTextSearchMatches,
            MaxTextSearchFlowComparisons = MaxTextSearchFlowComparisons,
            MaxTextSearchTableDetectionWork = MaxTextSearchTableDetectionWork,
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
            MaxFunctionShadingPixels = MaxFunctionShadingPixels,
            MaxFunctionShadingEvaluationWork = MaxFunctionShadingEvaluationWork,
            MaxContentOperands = MaxContentOperands,
            MaxContentNestingDepth = MaxContentNestingDepth,
            MaxType3GlyphInvocationsPerPage = MaxType3GlyphInvocationsPerPage,
            MaxImageReferenceSteps = MaxImageReferenceSteps,
            MaxPositionedTextCharactersPerPage = MaxPositionedTextCharactersPerPage,
            MaxPositionedTextWorkCharactersPerPage = MaxPositionedTextWorkCharactersPerPage,
            MaxPositionedTextProjectionCharactersPerPage = MaxPositionedTextProjectionCharactersPerPage,
            MaxClippedTextFontCopyWorkPerPage = MaxClippedTextFontCopyWorkPerPage,
            MaxPrintProductionContexts = MaxPrintProductionContexts,
            MaxPrintProductionOperations = MaxPrintProductionOperations,
        };
        snapshot.Validate();
        return snapshot;
    }
}
