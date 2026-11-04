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
        internal WordSaveOptions Clone() => (WordSaveOptions)MemberwiseClone();
        /// <summary>
        /// Controls saves of documents projected from legacy DOC files when known legacy-only
        /// content cannot be represented by the selected output format. The default blocks the save.
        /// </summary>
        public bool AllowLoss { get; set; }

        internal OfficeConversionLossPolicy LossPolicy {
            get => AllowLoss ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block;
            set => AllowLoss = value == OfficeConversionLossPolicy.Allow;
        }

        /// <summary>Physical DOCX or DOC format for streams and byte output. Paths select their native format by extension.</summary>
        public WordFileFormat Format { get; set; } = WordFileFormat.Docx;

        private string? _password;
        private bool _passwordSpecified;
        /// <summary>Output password. An explicitly assigned null removes encryption; an unset value preserves loaded encryption.</summary>
        public string? Password {
            get => _password;
            set { _password = value; _passwordSpecified = true; }
        }
        internal string? ResolvePassword(string? loadedPassword) => _passwordSpecified ? _password : loadedPassword;

        /// <summary>
        /// Gets or sets how save operations handle documents carrying digital-signature metadata.
        /// </summary>
        public WordSignedDocumentSavePolicy SignedDocumentPolicy { get; set; } = WordSignedDocumentSavePolicy.Block;

        /// <summary>
        /// Returns an options instance with default save behavior.
        /// </summary>
        public static WordSaveOptions Default => new();

        internal WordSaveOptions WithLossPolicy(OfficeConversionLossPolicy lossPolicy) {
            return new WordSaveOptions {
                LossPolicy = lossPolicy,
                SignedDocumentPolicy = SignedDocumentPolicy,
                Format = Format,
                _password = _password,
                _passwordSpecified = _passwordSpecified
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
