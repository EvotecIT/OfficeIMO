namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>
    ///     Represents an ActiveX control and its exact Office Forms storage.
    /// </summary>
    public sealed class LegacyPptActiveXControl {
        private readonly byte[] _storageBytes;
        private readonly byte[]? _metafileRecordBytes;

        internal LegacyPptActiveXControl(uint id, uint persistId,
            uint slideId, LegacyPptOleDrawAspect drawAspect, uint subType,
            string? menuName, string? progId, string? clipboardName,
            bool wasCompressed, byte[]? metafileRecordBytes,
            byte[] storageBytes) {
            Id = id;
            PersistId = persistId;
            SlideId = slideId;
            DrawAspect = drawAspect;
            SubType = subType;
            MenuName = menuName;
            ProgId = progId;
            ClipboardName = clipboardName;
            WasCompressed = wasCompressed;
            _metafileRecordBytes = metafileRecordBytes == null
                ? null
                : (byte[])metafileRecordBytes.Clone();
            _storageBytes = (byte[])(storageBytes
                ?? throw new ArgumentNullException(nameof(storageBytes)))
                .Clone();
        }

        /// <summary>Gets the document-wide external-object identifier.</summary>
        public uint Id { get; }
        /// <summary>Gets the control-storage persist identifier.</summary>
        public uint PersistId { get; }
        /// <summary>Gets the associated slide identifier, or zero.</summary>
        public uint SlideId { get; }
        /// <summary>Gets the view used to display the control.</summary>
        public LegacyPptOleDrawAspect DrawAspect { get; }
        /// <summary>Gets the raw ExOleObjSubTypeEnum value.</summary>
        public uint SubType { get; }
        /// <summary>Gets the short UI name.</summary>
        public string? MenuName { get; }
        /// <summary>Gets the ActiveX programmatic class identifier.</summary>
        public string? ProgId { get; }
        /// <summary>Gets the descriptive clipboard class name.</summary>
        public string? ClipboardName { get; }
        /// <summary>Gets whether the control storage was compressed.</summary>
        public bool WasCompressed { get; }
        /// <summary>Gets whether an optional icon metafile record is present.</summary>
        public bool HasMetafile => _metafileRecordBytes != null;
        /// <summary>Gets the optional icon metafile record byte count.</summary>
        public int MetafileByteCount => _metafileRecordBytes?.Length ?? 0;
        /// <summary>Gets the decoded Office Forms storage byte length.</summary>
        public int Length => _storageBytes.Length;
        /// <summary>Returns a defensive copy of the complete optional metafile record.</summary>
        public byte[]? GetMetafileRecordBytes() => _metafileRecordBytes == null
            ? null
            : (byte[])_metafileRecordBytes.Clone();
        /// <summary>Returns a defensive copy of the control storage.</summary>
        public byte[] GetBytes() => (byte[])_storageBytes.Clone();
    }
}
