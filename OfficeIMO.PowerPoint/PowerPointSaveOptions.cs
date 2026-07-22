namespace OfficeIMO.PowerPoint {
    /// <summary>Controls whether conversion to a less expressive PowerPoint format can omit content.</summary>
    public enum PowerPointConversionLossPolicy {
        /// <summary>Reject a conversion when known content or formatting cannot be represented.</summary>
        Block,

        /// <summary>Allow a conversion after known losses have been reported by preflight.</summary>
        Allow
    }

    /// <summary>Controls PowerPoint save and conversion behavior.</summary>
    public sealed class PowerPointSaveOptions {
        /// <summary>Gets or sets how known conversion loss is handled.</summary>
        public PowerPointConversionLossPolicy LossPolicy { get; set; } = PowerPointConversionLossPolicy.Block;

        /// <summary>
        /// Gets or sets the RC4 CryptoAPI key size used by encrypted PPT/POT/PPS saves.
        /// Valid values are 40 through 128 bits in 8-bit increments; the default is 128.
        /// </summary>
        public int LegacyPptEncryptionKeySizeBits { get; set; } = 128;

        /// <summary>
        /// Gets or sets whether encrypted PPT/POT/PPS saves also encrypt the Office
        /// document-property streams. The default is <see langword="true"/>.
        /// </summary>
        public bool LegacyPptEncryptDocumentProperties { get; set; } = true;

        /// <summary>
        /// Gets or sets whether legacy encrypted saves may retain compound
        /// streams that the RC4 CryptoAPI format does not encrypt. The default
        /// is <see langword="false"/> so password-protected output cannot
        /// silently contain unknown clear-text payloads.
        /// </summary>
        public bool LegacyPptAllowUnencryptedCompoundStreams { get; set; }
    }
}
