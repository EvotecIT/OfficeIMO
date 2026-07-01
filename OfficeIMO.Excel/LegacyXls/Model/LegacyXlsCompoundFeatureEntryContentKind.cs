namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Classifies the preserve-only content shape of a compound directory entry.
    /// </summary>
    public enum LegacyXlsCompoundFeatureEntryContentKind {
        /// <summary>The entry content could not be classified.</summary>
        Unknown,

        /// <summary>The entry is a compound storage, not a stream.</summary>
        Storage,

        /// <summary>The entry is an empty stream.</summary>
        EmptyStream,

        /// <summary>The entry is a VBA compressed container stream.</summary>
        VbaCompressedContainer,

        /// <summary>The entry is a VBA project metadata stream.</summary>
        VbaProjectMetadataStream,

        /// <summary>The entry is an embedded OLE stream payload.</summary>
        OlePayloadStream,

        /// <summary>The entry is a digital signature payload stream.</summary>
        DigitalSignatureStream,

        /// <summary>The entry is a stream with bytes that are not modeled more specifically yet.</summary>
        BinaryStream
    }
}
