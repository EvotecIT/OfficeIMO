namespace OfficeIMO.Word {
    internal enum WordFieldEvaluationReason {
        Default,
        ContainingResultReplaced,
        ExternalLayoutRequired,
        LocaleProfileUnsupported,
        ComplexTableProfileUnsupported,
        NestedInstruction,
        ListNumberingProfileUnsupported,
        CallerProvidedDateTime,
        RuntimeClock
    }

    /// <summary>Describes the evidence OfficeIMO used to produce or defer a field result.</summary>
    public enum WordFieldEvaluationBasis {
        /// <summary>The field was not evaluated.</summary>
        NotEvaluated,
        /// <summary>The result was derived deterministically from package content using invariant formatting rules.</summary>
        InvariantDocumentModel,
        /// <summary>The result was estimated from explicit page and section breaks, without a pagination engine.</summary>
        ExplicitBreakEstimate,
        /// <summary>The result requires Word or another layout-aware application.</summary>
        ExternalLayoutRequired,
        /// <summary>The result used the explicit update date/time supplied by the caller.</summary>
        CallerProvidedDateTime,
        /// <summary>The result used the host runtime clock because no update date/time was supplied by the caller.</summary>
        RuntimeClock
    }

    /// <summary>Describes how generated TOC, index, and caption-list page references were calculated.</summary>
    public enum WordPageNumberBasis {
        /// <summary>Page references were estimated from explicit page and section breaks.</summary>
        ExplicitBreakEstimate
    }

    internal static class WordFieldEvaluationContracts {
        internal static WordFieldEvaluationBasis GetBasis(
            WordFieldType? fieldType,
            WordFieldUpdateStatus status,
            bool isLocked,
            WordFieldEvaluationReason reason) {
            if (reason == WordFieldEvaluationReason.ContainingResultReplaced) {
                return WordFieldEvaluationBasis.NotEvaluated;
            }

            if (reason == WordFieldEvaluationReason.ExternalLayoutRequired) {
                return WordFieldEvaluationBasis.ExternalLayoutRequired;
            }

            if (status == WordFieldUpdateStatus.Updated && reason == WordFieldEvaluationReason.RuntimeClock) {
                return WordFieldEvaluationBasis.RuntimeClock;
            }

            if (status == WordFieldUpdateStatus.Updated && reason == WordFieldEvaluationReason.CallerProvidedDateTime) {
                return WordFieldEvaluationBasis.CallerProvidedDateTime;
            }

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

        internal static string GetDiagnosticCode(
            WordFieldType? fieldType,
            WordFieldUpdateStatus status,
            bool isLocked,
            WordFieldEvaluationReason reason) {
            if (status == WordFieldUpdateStatus.ParseError) return "FieldInstructionInvalid";
            if (reason == WordFieldEvaluationReason.ContainingResultReplaced) return "FieldContainingResultReplaced";
            if (isLocked) return "FieldLocked";
            if (status == WordFieldUpdateStatus.Skipped && fieldType is WordFieldType.TOC or WordFieldType.Index) return "FieldRefreshDelegated";
            if (reason == WordFieldEvaluationReason.ExternalLayoutRequired) return "FieldExternalLayoutRequired";
            if (reason == WordFieldEvaluationReason.LocaleProfileUnsupported) return "FieldLocaleProfileUnsupported";
            if (reason == WordFieldEvaluationReason.ComplexTableProfileUnsupported) return "FieldComplexTableProfileUnsupported";
            if (reason == WordFieldEvaluationReason.NestedInstruction) return "FieldNestedInstructionUnsupported";
            if (reason == WordFieldEvaluationReason.ListNumberingProfileUnsupported) return "FieldListNumberingProfileUnsupported";
            if (status == WordFieldUpdateStatus.Unsupported) return "FieldEvaluationUnsupported";
            if (status == WordFieldUpdateStatus.Skipped) return "FieldSourceUnavailable";
            if (reason == WordFieldEvaluationReason.RuntimeClock) return "FieldUpdatedFromRuntimeClock";
            if (reason == WordFieldEvaluationReason.CallerProvidedDateTime) return "FieldUpdatedFromCallerDateTime";
            return GetBasis(fieldType, status, isLocked, reason) == WordFieldEvaluationBasis.ExplicitBreakEstimate
                ? "FieldUpdatedFromExplicitBreaks"
                : "FieldUpdatedInvariant";
        }
    }
}
