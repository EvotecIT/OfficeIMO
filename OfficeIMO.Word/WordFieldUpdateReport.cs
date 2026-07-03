namespace OfficeIMO.Word {
    /// <summary>
    /// Describes the outcome of an OfficeIMO field update attempt.
    /// </summary>
    public enum WordFieldUpdateStatus {
        /// <summary>The field result was updated.</summary>
        Updated,
        /// <summary>The field was recognized but intentionally left unchanged.</summary>
        Skipped,
        /// <summary>The field type is not evaluated by OfficeIMO.</summary>
        Unsupported,
        /// <summary>The field instruction could not be parsed.</summary>
        ParseError
    }

    /// <summary>
    /// Describes a single field update outcome.
    /// </summary>
    public sealed class WordFieldUpdateResult {
        internal WordFieldUpdateResult(
            int index,
            WordFieldRepresentation representation,
            WordFieldLocationKind locationKind,
            string partUri,
            string instructionText,
            WordFieldType? fieldType,
            WordFieldUpdateStatus status,
            string? resultText,
            string message) {
            Index = index;
            Representation = representation;
            LocationKind = locationKind;
            PartUri = partUri;
            InstructionText = instructionText;
            FieldType = fieldType;
            Status = status;
            ResultText = resultText;
            Message = message;
        }

        /// <summary>Gets the deterministic index of the field in document scan order.</summary>
        public int Index { get; }

        /// <summary>Gets how the field is represented in the Open XML document.</summary>
        public WordFieldRepresentation Representation { get; }

        /// <summary>Gets the owning document part category.</summary>
        public WordFieldLocationKind LocationKind { get; }

        /// <summary>Gets the package part URI that contains the field.</summary>
        public string PartUri { get; }

        /// <summary>Gets the original field instruction text.</summary>
        public string InstructionText { get; }

        /// <summary>Gets the parsed field type when available.</summary>
        public WordFieldType? FieldType { get; }

        /// <summary>Gets the outcome of the update attempt.</summary>
        public WordFieldUpdateStatus Status { get; }

        /// <summary>Gets the result text written to the field, when one was written.</summary>
        public string? ResultText { get; }

        /// <summary>Gets a short diagnostic message for the outcome.</summary>
        public string Message { get; }
    }

    /// <summary>
    /// Summarizes a deterministic field update pass.
    /// </summary>
    public sealed class WordFieldUpdateReport {
        internal WordFieldUpdateReport(IReadOnlyList<WordFieldUpdateResult> results) {
            Results = results.ToArray();
        }

        /// <summary>Gets all field update results in document scan order.</summary>
        public IReadOnlyList<WordFieldUpdateResult> Results { get; }

        /// <summary>Gets the number of field results that were updated.</summary>
        public int UpdatedCount => Results.Count(result => result.Status == WordFieldUpdateStatus.Updated);

        /// <summary>Gets the number of recognized fields that were left unchanged.</summary>
        public int SkippedCount => Results.Count(result => result.Status == WordFieldUpdateStatus.Skipped);

        /// <summary>Gets the number of unsupported fields.</summary>
        public int UnsupportedCount => Results.Count(result => result.Status == WordFieldUpdateStatus.Unsupported);

        /// <summary>Gets the number of fields with parse errors.</summary>
        public int ParseErrorCount => Results.Count(result => result.Status == WordFieldUpdateStatus.ParseError);

        /// <summary>Gets the total number of fields inspected during the update pass.</summary>
        public int TotalCount => Results.Count;
    }
}
