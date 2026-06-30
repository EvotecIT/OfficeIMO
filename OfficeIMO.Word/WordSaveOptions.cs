namespace OfficeIMO.Word {
    /// <summary>
    /// Selects the physical document format for stream saves, where no file extension is available.
    /// </summary>
    public enum WordStreamSaveFormat {
        /// <summary>
        /// Save streams as the standard Office Open XML package format.
        /// </summary>
        OpenXml = 0,

        /// <summary>
        /// Save streams as a native Word 97-2003 legacy .doc compound file.
        /// </summary>
        LegacyDoc = 1
    }

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
        /// <summary>
        /// Selects the physical document format for <see cref="WordDocument.Save(System.IO.Stream, WordSaveOptions?)"/>.
        /// File-path saves continue to use the destination extension.
        /// </summary>
        public WordStreamSaveFormat StreamFormat { get; set; }

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
