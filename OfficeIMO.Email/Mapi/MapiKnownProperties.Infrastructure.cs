namespace OfficeIMO.Email;

public static partial class MapiKnownProperties {
    public static partial class PidTag {
        /// <summary>PidTagParentEntryId (0x0E09).</summary>
        public static readonly MapiPropertyKey<byte[]> ParentEntryId = Binary("PidTagParentEntryId", 0x0E09);
        /// <summary>PidTagExtendedFolderFlags (0x36DA).</summary>
        public static readonly MapiPropertyKey<byte[]> ExtendedFolderFlags =
            Binary("PidTagExtendedFolderFlags", 0x36DA);
        /// <summary>PidTagAccess (0x0FF4).</summary>
        public static readonly MapiPropertyKey<int> Access = Integer("PidTagAccess", 0x0FF4);
        /// <summary>PidTagRowType (0x0FF5).</summary>
        public static readonly MapiPropertyKey<int> RowType = Integer("PidTagRowType", 0x0FF5);
        /// <summary>PidTagInstanceKey (0x0FF6).</summary>
        public static readonly MapiPropertyKey<byte[]> InstanceKey = Binary("PidTagInstanceKey", 0x0FF6);
        /// <summary>PidTagAccessLevel (0x0FF7).</summary>
        public static readonly MapiPropertyKey<int> AccessLevel = Integer("PidTagAccessLevel", 0x0FF7);
        /// <summary>PidTagMappingSignature (0x0FF8).</summary>
        public static readonly MapiPropertyKey<byte[]> MappingSignature = Binary("PidTagMappingSignature", 0x0FF8);
        /// <summary>PidTagStoreRecordKey (0x0FFA).</summary>
        public static readonly MapiPropertyKey<byte[]> StoreRecordKey = Binary("PidTagStoreRecordKey", 0x0FFA);
        /// <summary>PidTagStoreEntryId (0x0FFB).</summary>
        public static readonly MapiPropertyKey<byte[]> StoreEntryId = Binary("PidTagStoreEntryId", 0x0FFB);
        /// <summary>PidTagSourceKey (0x65E0).</summary>
        public static readonly MapiPropertyKey<byte[]> SourceKey = Binary("PidTagSourceKey", 0x65E0);
        /// <summary>PidTagParentSourceKey (0x65E1).</summary>
        public static readonly MapiPropertyKey<byte[]> ParentSourceKey = Binary("PidTagParentSourceKey", 0x65E1);
        /// <summary>PidTagChangeKey (0x65E2).</summary>
        public static readonly MapiPropertyKey<byte[]> ChangeKey = Binary("PidTagChangeKey", 0x65E2);
        /// <summary>PidTagPredecessorChangeList (0x65E3).</summary>
        public static readonly MapiPropertyKey<byte[]> PredecessorChangeList =
            Binary("PidTagPredecessorChangeList", 0x65E3);
        /// <summary>PidTagChangeNumber (0x67A4).</summary>
        public static readonly MapiPropertyKey<long> ChangeNumber = new MapiPropertyKey<long>(
            "PidTagChangeNumber", 0x67A4, MapiPropertyType.Integer64);

        internal static readonly IReadOnlyList<MapiPropertyKey> InfrastructureTaggedProperties =
            new MapiPropertyKey[] {
                ParentEntryId, ExtendedFolderFlags, Access, RowType, InstanceKey, AccessLevel, MappingSignature, StoreRecordKey,
                StoreEntryId, SourceKey, ParentSourceKey, ChangeKey, PredecessorChangeList, ChangeNumber
            };
    }
}
