namespace OfficeIMO.Email;

public static partial class MapiKnownProperties {
    /// <summary>Tagged properties used by PST/OST store and table structures.</summary>
    public static partial class PidTag {
        /// <summary>PidTagMessageToMe (0x0057).</summary>
        public static readonly MapiPropertyKey<bool> MessageToMe = Boolean("PidTagMessageToMe", 0x0057);
        /// <summary>PidTagMessageCcMe (0x0058).</summary>
        public static readonly MapiPropertyKey<bool> MessageCcMe = Boolean("PidTagMessageCcMe", 0x0058);
        /// <summary>PidTagReplItemid (0x0E30).</summary>
        public static readonly MapiPropertyKey<int> ReplItemid = Integer("PidTagReplItemid", 0x0E30);
        /// <summary>PidTagReplChangenum (0x0E33).</summary>
        public static readonly MapiPropertyKey<long> ReplChangenum = new MapiPropertyKey<long>(
            "PidTagReplChangenum", 0x0E33, MapiPropertyType.Integer64);
        /// <summary>PidTagReplVersionHistory (0x0E34).</summary>
        public static readonly MapiPropertyKey<byte[]> ReplVersionHistory = Binary("PidTagReplVersionHistory", 0x0E34);
        /// <summary>PidTagReplFlags (0x0E38).</summary>
        public static readonly MapiPropertyKey<int> ReplFlags = Integer("PidTagReplFlags", 0x0E38);
        /// <summary>PidTagReplCopiedfromVersionhistory (0x0E3C).</summary>
        public static readonly MapiPropertyKey<byte[]> ReplCopiedfromVersionhistory =
            Binary("PidTagReplCopiedfromVersionhistory", 0x0E3C);
        /// <summary>PidTagReplCopiedfromItemid (0x0E3D).</summary>
        public static readonly MapiPropertyKey<byte[]> ReplCopiedfromItemid =
            Binary("PidTagReplCopiedfromItemid", 0x0E3D);
        /// <summary>PidTagValidFolderMask (0x35DF).</summary>
        public static readonly MapiPropertyKey<int> ValidFolderMask = Integer("PidTagValidFolderMask", 0x35DF);
        /// <summary>PidTagIpmSubTreeEntryId (0x35E0).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmSubTreeEntryId = Binary("PidTagIpmSubTreeEntryId", 0x35E0);
        /// <summary>PidTagIpmInboxEntryId (0x35E1).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmInboxEntryId = Binary("PidTagIpmInboxEntryId", 0x35E1);
        /// <summary>PidTagIpmOutboxEntryId (0x35E2).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmOutboxEntryId = Binary("PidTagIpmOutboxEntryId", 0x35E2);
        /// <summary>PidTagIpmWastebasketEntryId (0x35E3).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmWastebasketEntryId =
            Binary("PidTagIpmWastebasketEntryId", 0x35E3);
        /// <summary>PidTagIpmSentMailEntryId (0x35E4).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmSentMailEntryId = Binary("PidTagIpmSentMailEntryId", 0x35E4);
        /// <summary>PidTagViewsEntryId (0x35E5).</summary>
        public static readonly MapiPropertyKey<byte[]> ViewsEntryId = Binary("PidTagViewsEntryId", 0x35E5);
        /// <summary>PidTagCommonViewsEntryId (0x35E6).</summary>
        public static readonly MapiPropertyKey<byte[]> CommonViewsEntryId = Binary("PidTagCommonViewsEntryId", 0x35E6);
        /// <summary>PidTagFinderEntryId (0x35E7).</summary>
        public static readonly MapiPropertyKey<byte[]> FinderEntryId = Binary("PidTagFinderEntryId", 0x35E7);
        /// <summary>PidTagIpmAppointmentEntryId (0x36D0).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmAppointmentEntryId =
            Binary("PidTagIpmAppointmentEntryId", 0x36D0);
        /// <summary>PidTagIpmContactEntryId (0x36D1).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmContactEntryId = Binary("PidTagIpmContactEntryId", 0x36D1);
        /// <summary>PidTagIpmJournalEntryId (0x36D2).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmJournalEntryId = Binary("PidTagIpmJournalEntryId", 0x36D2);
        /// <summary>PidTagIpmNoteEntryId (0x36D3).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmNoteEntryId = Binary("PidTagIpmNoteEntryId", 0x36D3);
        /// <summary>PidTagIpmTaskEntryId (0x36D4).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmTaskEntryId = Binary("PidTagIpmTaskEntryId", 0x36D4);
        /// <summary>PidTagIpmDraftsEntryId (0x36D7).</summary>
        public static readonly MapiPropertyKey<byte[]> IpmDraftsEntryId = Binary("PidTagIpmDraftsEntryId", 0x36D7);
        /// <summary>PidTagSecureSubmitFlags (0x65C6).</summary>
        public static readonly MapiPropertyKey<int> SecureSubmitFlags = Integer("PidTagSecureSubmitFlags", 0x65C6);
        /// <summary>PST search-criteria flags (0x660B).</summary>
        public static readonly MapiPropertyKey<int> PstSearchCriteriaFlags =
            Integer("PstSearchCriteriaFlags", 0x660B);
        /// <summary>PidTagPstHiddenCount (0x6635).</summary>
        public static readonly MapiPropertyKey<int> PstHiddenCount = Integer("PidTagPstHiddenCount", 0x6635);
        /// <summary>PidTagPstHiddenUnread (0x6636).</summary>
        public static readonly MapiPropertyKey<int> PstHiddenUnread = Integer("PidTagPstHiddenUnread", 0x6636);
        /// <summary>PidTagOfflineAddressBookName (0x6800).</summary>
        public static readonly MapiPropertyKey<string> OfflineAddressBookName =
            String("PidTagOfflineAddressBookName", 0x6800);
        /// <summary>PidTagRwRulesStream (0x6802).</summary>
        public static readonly MapiPropertyKey<byte[]> RwRulesStream = Binary("PidTagRwRulesStream", 0x6802);
        /// <summary>PidTagSendOutlookRecallReport (0x6803).</summary>
        public static readonly MapiPropertyKey<bool> SendOutlookRecallReport =
            Boolean("PidTagSendOutlookRecallReport", 0x6803);
        /// <summary>PidTagOfflineAddressBookTruncatedProperties (0x6805).</summary>
        public static readonly MapiPropertyKey<object> OfflineAddressBookTruncatedProperties =
            new MapiPropertyKey<object>("PidTagOfflineAddressBookTruncatedProperties", 0x6805,
                MapiPropertyType.MultipleInteger32);
        /// <summary>PidTagSearchFolderLastUsed (0x6834), in minutes since 1601-01-01 UTC.</summary>
        public static readonly MapiPropertyKey<int> SearchFolderLastUsed =
            Integer("PidTagSearchFolderLastUsed", 0x6834);
        /// <summary>PidTagSearchFolderExpiration (0x683A), in minutes since 1601-01-01 UTC.</summary>
        public static readonly MapiPropertyKey<int> SearchFolderExpiration =
            Integer("PidTagSearchFolderExpiration", 0x683A);
        /// <summary>PidTagSearchFolderTemplateId (0x6841).</summary>
        public static readonly MapiPropertyKey<int> SearchFolderTemplateId =
            Integer("PidTagSearchFolderTemplateId", 0x6841);
        /// <summary>PidTagSearchFolderId (0x6842).</summary>
        public static readonly MapiPropertyKey<byte[]> SearchFolderId = Binary("PidTagSearchFolderId", 0x6842);
        /// <summary>PidTagSearchFolderDefinition (0x6845).</summary>
        public static readonly MapiPropertyKey<byte[]> SearchFolderDefinition =
            Binary("PidTagSearchFolderDefinition", 0x6845);
        /// <summary>PidTagSearchFolderStorageType (0x6846).</summary>
        public static readonly MapiPropertyKey<int> SearchFolderStorageType =
            Integer("PidTagSearchFolderStorageType", 0x6846);
        /// <summary>PidTagSearchFolderTag (0x6847).</summary>
        public static readonly MapiPropertyKey<int> SearchFolderTag = Integer("PidTagSearchFolderTag", 0x6847);
        /// <summary>PidTagSearchFolderEfpFlags (0x6848).</summary>
        public static readonly MapiPropertyKey<int> SearchFolderEfpFlags =
            Integer("PidTagSearchFolderEfpFlags", 0x6848);
        /// <summary>PidTagViewDescriptorBinary (0x7001).</summary>
        public static readonly MapiPropertyKey<byte[]> ViewDescriptorBinary =
            Binary("PidTagViewDescriptorBinary", 0x7001);
        /// <summary>PidTagViewDescriptorStrings (0x7002).</summary>
        public static readonly MapiPropertyKey<byte[]> ViewDescriptorStrings =
            Binary("PidTagViewDescriptorStrings", 0x7002);
        /// <summary>PidTagViewDescriptorFlags (0x7003).</summary>
        public static readonly MapiPropertyKey<int> ViewDescriptorFlags = Integer("PidTagViewDescriptorFlags", 0x7003);
        /// <summary>PidTagViewDescriptorLinkTo (0x7004).</summary>
        public static readonly MapiPropertyKey<byte[]> ViewDescriptorLinkTo = Binary("PidTagViewDescriptorLinkTo", 0x7004);
        /// <summary>PidTagViewDescriptorViewFolder (0x7005).</summary>
        public static readonly MapiPropertyKey<byte[]> ViewDescriptorViewFolder =
            Binary("PidTagViewDescriptorViewFolder", 0x7005);
        /// <summary>PidTagViewDescriptorName (0x7006).</summary>
        public static readonly MapiPropertyKey<string> ViewDescriptorName =
            String("PidTagViewDescriptorName", 0x7006);
        /// <summary>PidTagViewDescriptorVersion (0x7007).</summary>
        public static readonly MapiPropertyKey<int> ViewDescriptorVersion =
            Integer("PidTagViewDescriptorVersion", 0x7007);
        /// <summary>PidTagRoamingDatatypes (0x7C06).</summary>
        public static readonly MapiPropertyKey<int> RoamingDatatypes = Integer("PidTagRoamingDatatypes", 0x7C06);
        /// <summary>PidTagRoamingDictionary (0x7C07).</summary>
        public static readonly MapiPropertyKey<byte[]> RoamingDictionary =
            Binary("PidTagRoamingDictionary", 0x7C07);
        /// <summary>PidTagRoamingXmlStream (0x7C08).</summary>
        public static readonly MapiPropertyKey<byte[]> RoamingXmlStream = Binary("PidTagRoamingXmlStream", 0x7C08);

        internal static readonly IReadOnlyList<MapiPropertyKey> PstAndTableProperties = new MapiPropertyKey[] {
            MessageToMe, MessageCcMe, ReplItemid, ReplChangenum, ReplVersionHistory, ReplFlags,
            ReplCopiedfromVersionhistory, ReplCopiedfromItemid, ValidFolderMask, IpmSubTreeEntryId,
            IpmInboxEntryId, IpmOutboxEntryId, IpmWastebasketEntryId, IpmSentMailEntryId, ViewsEntryId,
            CommonViewsEntryId, FinderEntryId, IpmAppointmentEntryId, IpmContactEntryId, IpmJournalEntryId,
            IpmNoteEntryId, IpmTaskEntryId, IpmDraftsEntryId, SecureSubmitFlags, PstSearchCriteriaFlags,
            PstHiddenCount, PstHiddenUnread, OfflineAddressBookName, RwRulesStream, SendOutlookRecallReport,
            OfflineAddressBookTruncatedProperties, SearchFolderLastUsed, SearchFolderExpiration,
            SearchFolderTemplateId, SearchFolderId, SearchFolderDefinition, SearchFolderStorageType,
            SearchFolderTag, SearchFolderEfpFlags, ViewDescriptorBinary, ViewDescriptorStrings,
            ViewDescriptorFlags, ViewDescriptorLinkTo, ViewDescriptorViewFolder, ViewDescriptorName,
            ViewDescriptorVersion, RoamingDatatypes, RoamingDictionary, RoamingXmlStream
        };
    }
}
