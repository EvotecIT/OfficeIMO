namespace OfficeIMO.Word {
    /// <summary>Describes the evidence OfficeIMO used to produce or defer a field result.</summary>
    public enum WordFieldEvaluationBasis {
        /// <summary>The field was not evaluated.</summary>
        NotEvaluated,
        /// <summary>The result was derived deterministically from package content using invariant formatting rules.</summary>
        InvariantDocumentModel,
        /// <summary>The result was estimated from explicit page and section breaks, without a pagination engine.</summary>
        ExplicitBreakEstimate,
        /// <summary>The result requires Word or another layout-aware application.</summary>
        ExternalLayoutRequired
    }

    /// <summary>Describes how generated TOC, index, and caption-list page references were calculated.</summary>
    public enum WordPageNumberBasis {
        /// <summary>Page references were estimated from explicit page and section breaks.</summary>
        ExplicitBreakEstimate
    }

    internal static class WordFieldEvaluationContracts {
        internal static WordFieldEvaluationBasis GetBasis(WordFieldType? fieldType, WordFieldUpdateStatus status, bool isLocked) {
            if (!isLocked && status == WordFieldUpdateStatus.Skipped && fieldType is WordFieldType.TOC or WordFieldType.Index) {
                return WordFieldEvaluationBasis.ExternalLayoutRequired;
            }

            if (status != WordFieldUpdateStatus.Updated) {
                return WordFieldEvaluationBasis.NotEvaluated;
            }

            if (fieldType is WordFieldType.TOC or WordFieldType.Index) {
                return WordFieldEvaluationBasis.ExternalLayoutRequired;
            }

            return fieldType is WordFieldType.Page or WordFieldType.PageRef or WordFieldType.NumPages or WordFieldType.SectionPages
                ? WordFieldEvaluationBasis.ExplicitBreakEstimate
                : WordFieldEvaluationBasis.InvariantDocumentModel;
        }

        internal static string GetDiagnosticCode(WordFieldType? fieldType, WordFieldUpdateStatus status, bool isLocked) {
            if (status == WordFieldUpdateStatus.ParseError) return "FieldInstructionInvalid";
            if (status == WordFieldUpdateStatus.Unsupported) return "FieldEvaluationUnsupported";
            if (isLocked) return "FieldLocked";
            if (status == WordFieldUpdateStatus.Skipped && fieldType is WordFieldType.TOC or WordFieldType.Index) return "FieldRefreshDelegated";
            if (status == WordFieldUpdateStatus.Skipped) return "FieldSourceUnavailable";
            return GetBasis(fieldType, status, isLocked) == WordFieldEvaluationBasis.ExplicitBreakEstimate
                ? "FieldUpdatedFromExplicitBreaks"
                : "FieldUpdatedInvariant";
        }
    }
}
