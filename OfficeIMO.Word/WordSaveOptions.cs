namespace OfficeIMO.Word {
    /// <summary>
    /// Describes how save operations should handle documents carrying digital-signature metadata.
    /// </summary>
    public enum WordSignedDocumentSavePolicy {
        /// <summary>
        /// Block saves when signature metadata is present because the resulting package may invalidate signatures.
        /// </summary>
        Block,

        /// <summary>
        /// Allow saving even though existing signatures may become invalid.
        /// </summary>
        AllowSignatureInvalidation
    }

    /// <summary>
    /// Optional behaviors applied during Word document save operations.
    /// </summary>
    public sealed class WordSaveOptions {
        /// <summary>Gets or sets whether to open the saved file after a successful file commit.</summary>
        public bool OpenAfterSave { get; set; }

        /// <summary>
        /// Controls saves of documents projected from legacy DOC files when known legacy-only
        /// content cannot be represented by the selected output format. The default blocks the save.
        /// </summary>
        public WordConversionLossPolicy LossPolicy { get; set; } = WordConversionLossPolicy.Block;

        /// <summary>
        /// Gets or sets how save operations handle documents carrying digital-signature metadata.
        /// </summary>
        public WordSignedDocumentSavePolicy SignedDocumentPolicy { get; set; } = WordSignedDocumentSavePolicy.Block;

        /// <summary>
        /// Returns an options instance with default save behavior.
        /// </summary>
        public static WordSaveOptions Default => new();

        /// <summary>
        /// Returns an options instance with all optional behaviors disabled.
        /// </summary>
        public static WordSaveOptions None => new();

        internal WordSaveOptions WithLossPolicy(WordConversionLossPolicy lossPolicy) {
            return new WordSaveOptions {
                OpenAfterSave = OpenAfterSave,
                LossPolicy = lossPolicy,
                SignedDocumentPolicy = SignedDocumentPolicy
            };
        }
    }

    /// <summary>
    /// Raised when a save operation is blocked because signature metadata is present.
    /// </summary>
    public sealed class WordSignatureSavePolicyException : InvalidOperationException {
        internal WordSignatureSavePolicyException(string operation, WordSignatureInfo signatureInfo)
            : base(CreateMessage(operation, signatureInfo)) {
            Operation = operation;
            SignatureInfo = signatureInfo;
        }

        /// <summary>
        /// Gets the save operation that was blocked.
        /// </summary>
        public string Operation { get; }

        /// <summary>
        /// Gets signature metadata discovered before the save was blocked.
        /// </summary>
        public WordSignatureInfo SignatureInfo { get; }

        private static string CreateMessage(string operation, WordSignatureInfo signatureInfo) {
            string count = signatureInfo.FindingCount.ToString(System.Globalization.CultureInfo.InvariantCulture);
            return operation + " was blocked because the document contains " + count
                + " digital-signature metadata item(s). Saving or rewriting a signed package may invalidate existing signatures. "
                + "Pass WordSaveOptions with SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation to save anyway.";
        }
    }
}
