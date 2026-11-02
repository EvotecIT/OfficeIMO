namespace OfficeIMO.PowerPoint {
    /// <summary>Controls PowerPoint save and conversion behavior.</summary>
    public sealed class PowerPointSaveOptions {
        internal PowerPointSaveOptions Clone() => (PowerPointSaveOptions)MemberwiseClone();
        /// <summary>Gets or sets how known conversion loss is handled.</summary>
        public bool AllowLoss { get; set; }
        internal OfficeConversionLossPolicy LossPolicy { get => AllowLoss ? OfficeConversionLossPolicy.Allow : OfficeConversionLossPolicy.Block; set => AllowLoss = value == OfficeConversionLossPolicy.Allow; }
        /// <summary>Physical output format for streams and byte encoding.</summary>
        public PowerPointFileFormat Format { get; set; } = PowerPointFileFormat.Pptx;
        private string? _password;
        private bool _passwordSpecified;
        private int _legacyKeySize = 128;
        private bool _legacyEncryptProperties = true, _legacyAllowClearStreams;
        private bool _legacyKeySizeSpecified, _legacyPropertiesSpecified, _legacyClearStreamsSpecified;
        internal bool HasPasswordOverride => _passwordSpecified;
        internal bool HasLegacyEncryptionOverride => _legacyKeySizeSpecified || _legacyPropertiesSpecified || _legacyClearStreamsSpecified;
        internal int ResolveLegacyKeySize(int? source) => _legacyKeySizeSpecified ? _legacyKeySize : source ?? 128;
        internal bool ResolveLegacyProperties(bool? source) => _legacyPropertiesSpecified ? _legacyEncryptProperties : source ?? true;
        /// <summary>Output password. An unset value preserves loaded encryption; explicit null removes it.</summary>
        public string? Password { get => _password; set { _password = value; _passwordSpecified = true; } }
        internal string? ResolvePassword(string? loaded) => _passwordSpecified ? _password : loaded;

        /// <summary>
        /// Gets or sets the RC4 CryptoAPI key size used by encrypted PPT/POT/PPS saves.
        /// Valid values are 40 through 128 bits in 8-bit increments; the default is 128.
        /// </summary>
        public int LegacyPptEncryptionKeySizeBits { get => _legacyKeySize; set { _legacyKeySize = value; _legacyKeySizeSpecified = true; } }

        /// <summary>
        /// Gets or sets whether encrypted PPT/POT/PPS saves also encrypt the Office
        /// document-property streams. The default is <see langword="true"/>.
        /// </summary>
        public bool LegacyPptEncryptDocumentProperties { get => _legacyEncryptProperties; set { _legacyEncryptProperties = value; _legacyPropertiesSpecified = true; } }

        /// <summary>
        /// Gets or sets whether legacy encrypted saves may retain compound
        /// streams that the RC4 CryptoAPI format does not encrypt. The default
        /// is <see langword="false"/> so password-protected output cannot
        /// silently contain unknown clear-text payloads.
        /// </summary>
        public bool LegacyPptAllowUnencryptedCompoundStreams { get => _legacyAllowClearStreams; set { _legacyAllowClearStreams = value; _legacyClearStreamsSpecified = true; } }
    }
}
