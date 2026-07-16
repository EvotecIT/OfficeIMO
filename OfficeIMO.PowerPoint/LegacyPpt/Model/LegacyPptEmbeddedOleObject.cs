namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Specifies the view used to display an embedded OLE object.</summary>
    public enum LegacyPptOleDrawAspect : uint {
        /// <summary>The object's content view.</summary>
        Content = 0x00000001,
        /// <summary>A thumbnail view.</summary>
        Thumbnail = 0x00000002,
        /// <summary>The object's icon.</summary>
        Icon = 0x00000004,
        /// <summary>The print view.</summary>
        DocumentPrint = 0x00000008
    }

    /// <summary>Specifies how an embedded object follows the presentation colors.</summary>
    public enum LegacyPptOleColorFollow : uint {
        /// <summary>The object does not follow the color scheme.</summary>
        None = 0,
        /// <summary>The object follows the full color scheme.</summary>
        Scheme = 1,
        /// <summary>The object follows text and background colors.</summary>
        TextAndBackground = 2
    }

    /// <summary>Represents one decoded embedded OLE compound storage.</summary>
    public sealed class LegacyPptEmbeddedOleObject {
        private readonly byte[] _storageBytes;

        internal LegacyPptEmbeddedOleObject(uint id, uint persistId,
            LegacyPptOleDrawAspect drawAspect, uint subType,
            LegacyPptOleColorFollow colorFollow, bool cannotLockServer,
            bool noSizeToServer, bool isTable, string? menuName,
            string? progId, string? clipboardName, bool wasCompressed,
            byte[] storageBytes) {
            Id = id;
            PersistId = persistId;
            DrawAspect = drawAspect;
            SubType = subType;
            ColorFollow = colorFollow;
            CannotLockServer = cannotLockServer;
            NoSizeToServer = noSizeToServer;
            IsTable = isTable;
            MenuName = menuName;
            ProgId = progId;
            ClipboardName = clipboardName;
            WasCompressed = wasCompressed;
            _storageBytes = (byte[])(storageBytes
                ?? throw new ArgumentNullException(nameof(storageBytes))).Clone();
        }

        /// <summary>Gets the document-wide external-object identifier.</summary>
        public uint Id { get; }

        /// <summary>Gets the persist identifier of the ExOleObjStg record.</summary>
        public uint PersistId { get; }

        /// <summary>Gets the view used to display the object.</summary>
        public LegacyPptOleDrawAspect DrawAspect { get; }

        /// <summary>Gets the raw ExOleObjSubTypeEnum value.</summary>
        public uint SubType { get; }

        /// <summary>Gets how the object follows the presentation colors.</summary>
        public LegacyPptOleColorFollow ColorFollow { get; }

        /// <summary>Gets whether the OLE server cannot be locked.</summary>
        public bool CannotLockServer { get; }

        /// <summary>Gets whether dimensions need not be sent to the OLE server.</summary>
        public bool NoSizeToServer { get; }

        /// <summary>Gets whether the object represents a Word table.</summary>
        public bool IsTable { get; }

        /// <summary>Gets the short UI name.</summary>
        public string? MenuName { get; }

        /// <summary>Gets the object's programmatic class identifier.</summary>
        public string? ProgId { get; }

        /// <summary>Gets the descriptive clipboard class name.</summary>
        public string? ClipboardName { get; }

        /// <summary>Gets whether the source persist record was compressed.</summary>
        public bool WasCompressed { get; }

        /// <summary>Gets the decoded compound-storage byte length.</summary>
        public int Length => _storageBytes.Length;

        /// <summary>Returns a defensive copy of the compound storage.</summary>
        public byte[] GetBytes() => (byte[])_storageBytes.Clone();
    }
}
