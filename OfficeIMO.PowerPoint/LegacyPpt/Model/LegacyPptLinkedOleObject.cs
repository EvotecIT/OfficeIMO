namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Specifies when a linked OLE cache is updated.</summary>
    public enum LegacyPptOleUpdateMode : uint {
        /// <summary>Update the linked cache whenever possible.</summary>
        Automatic = 1,
        /// <summary>Update the linked cache only on request.</summary>
        Manual = 3
    }

    /// <summary>
    ///     Represents a linked OLE object and its exact native cache storage.
    /// </summary>
    public sealed class LegacyPptLinkedOleObject {
        private readonly byte[] _storageBytes;
        private readonly byte[]? _metafileRecordBytes;

        internal LegacyPptLinkedOleObject(uint id, uint persistId,
            uint slideId, LegacyPptOleUpdateMode updateMode,
            LegacyPptOleDrawAspect drawAspect, uint subType,
            string? menuName, string? progId, string? clipboardName,
            bool wasCompressed, byte[]? metafileRecordBytes,
            byte[] storageBytes) {
            Id = id;
            PersistId = persistId;
            SlideId = slideId;
            UpdateMode = updateMode;
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
        /// <summary>Gets the cache-storage persist identifier.</summary>
        public uint PersistId { get; }
        /// <summary>Gets the associated slide identifier, or zero.</summary>
        public uint SlideId { get; }
        /// <summary>Gets when the linked cache is updated.</summary>
        public LegacyPptOleUpdateMode UpdateMode { get; }
        /// <summary>Gets the view used to display the object.</summary>
        public LegacyPptOleDrawAspect DrawAspect { get; }
        /// <summary>Gets the raw ExOleObjSubTypeEnum value.</summary>
        public uint SubType { get; }
        /// <summary>Gets the short UI name.</summary>
        public string? MenuName { get; }
        /// <summary>Gets the programmatic class identifier.</summary>
        public string? ProgId { get; }
        /// <summary>Gets the descriptive clipboard class name.</summary>
        public string? ClipboardName { get; }
        /// <summary>Gets whether the cache used compressed storage.</summary>
        public bool WasCompressed { get; }
        /// <summary>Gets whether an optional icon metafile record is present.</summary>
        public bool HasMetafile => _metafileRecordBytes != null;
        /// <summary>Gets the optional icon metafile record byte count.</summary>
        public int MetafileByteCount => _metafileRecordBytes?.Length ?? 0;
        /// <summary>Gets the decoded cache-storage byte length.</summary>
        public int Length => _storageBytes.Length;
        /// <summary>Returns a defensive copy of the complete optional metafile record.</summary>
        public byte[]? GetMetafileRecordBytes() => _metafileRecordBytes == null
            ? null
            : (byte[])_metafileRecordBytes.Clone();
        /// <summary>Returns a defensive copy of the cache storage.</summary>
        public byte[] GetBytes() => (byte[])_storageBytes.Clone();
    }
}
