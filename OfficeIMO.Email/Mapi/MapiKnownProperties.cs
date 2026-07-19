namespace OfficeIMO.Email;

/// <summary>
/// Typed vocabulary for well-known tagged, numeric named, and string named MAPI properties supported by OfficeIMO.
/// </summary>
public static partial class MapiKnownProperties {
    // Lazy construction avoids a .NET Framework static-initialization cycle when a nested vocabulary class calls
    // one of the key factories on this containing type while its own field table is still being initialized.
    private static readonly Lazy<IReadOnlyList<MapiPropertyKey>> Known =
        new Lazy<IReadOnlyList<MapiPropertyKey>>(() => PidTag.All
            .Concat(PidLid.All)
            .Concat(PidName.All)
            .ToArray(), LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>All property keys currently published by OfficeIMO.</summary>
    public static IReadOnlyList<MapiPropertyKey> All => Known.Value;

    /// <summary>Finds a published key by canonical name.</summary>
    public static MapiPropertyKey? Find(string canonicalName) {
        if (canonicalName == null) throw new ArgumentNullException(nameof(canonicalName));
        return Known.Value.FirstOrDefault(key => string.Equals(key.CanonicalName, canonicalName,
            StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Finds a published standard-property key by property identifier.</summary>
    public static MapiPropertyKey? Find(ushort propertyId) {
        MapiPropertyKey[] matches = Known.Value.Where(key => !key.IsNamed && key.PropertyId == propertyId).ToArray();
        return matches.Length == 1 ? matches[0] : null;
    }

    /// <summary>Finds a published standard-property key by property identifier and wire type.</summary>
    public static MapiPropertyKey? Find(ushort propertyId, MapiPropertyType propertyType) {
        return Known.Value.FirstOrDefault(key => !key.IsNamed && key.PropertyId == propertyId && key.Accepts(propertyType));
    }

    /// <summary>Finds a published key matching a retained MAPI property's identity and wire type.</summary>
    public static MapiPropertyKey? Find(MapiProperty property) {
        if (property == null) throw new ArgumentNullException(nameof(property));
        return Known.Value.FirstOrDefault(key => key.Matches(property));
    }

    /// <summary>Finds a published numeric named-property key.</summary>
    public static MapiPropertyKey? Find(Guid propertySet, uint localId) {
        return Known.Value.FirstOrDefault(key => key.Name?.PropertySet == propertySet && key.Name.LocalId == localId);
    }

    /// <summary>Finds a published string named-property key.</summary>
    public static MapiPropertyKey? Find(Guid propertySet, string name) {
        if (name == null) throw new ArgumentNullException(nameof(name));
        return Known.Value.FirstOrDefault(key => key.Name?.PropertySet == propertySet &&
            string.Equals(key.Name.Name, name, StringComparison.OrdinalIgnoreCase));
    }

    private static MapiPropertyKey<string> String(string name, ushort id) =>
        new MapiPropertyKey<string>(name, id, MapiPropertyType.Unicode, MapiPropertyType.String8);

    private static MapiPropertyKey<int> Integer(string name, ushort id) =>
        new MapiPropertyKey<int>(name, id, MapiPropertyType.Integer32,
            MapiPropertyType.Integer16, MapiPropertyType.Integer64);

    private static MapiPropertyKey<bool> Boolean(string name, ushort id) =>
        new MapiPropertyKey<bool>(name, id, MapiPropertyType.Boolean);

    private static MapiPropertyKey<DateTimeOffset> Time(string name, ushort id) =>
        new MapiPropertyKey<DateTimeOffset>(name, id, MapiPropertyType.Time);

    private static MapiPropertyKey<byte[]> Binary(string name, ushort id) =>
        new MapiPropertyKey<byte[]>(name, id, MapiPropertyType.Binary);

    /// <summary>Well-known tagged properties (PidTag).</summary>
    public static partial class PidTag {
        /// <summary>PidTagNameidBucketCount (0x0001).</summary>
        public static readonly MapiPropertyKey<int> NameidBucketCount = Integer("PidTagNameidBucketCount", 0x0001);
        /// <summary>PidTagNameidStreamGuid (0x0002).</summary>
        public static readonly MapiPropertyKey<byte[]> NameidStreamGuid = Binary("PidTagNameidStreamGuid", 0x0002);
        /// <summary>PidTagNameidStreamEntry (0x0003).</summary>
        public static readonly MapiPropertyKey<byte[]> NameidStreamEntry = Binary("PidTagNameidStreamEntry", 0x0003);
        /// <summary>PidTagNameidStreamString (0x0004).</summary>
        public static readonly MapiPropertyKey<byte[]> NameidStreamString = Binary("PidTagNameidStreamString", 0x0004);
        /// <summary>PidTagAlternateRecipientAllowed (0x0002 on message objects).</summary>
        public static readonly MapiPropertyKey<bool> AlternateRecipientAllowed =
            Boolean("PidTagAlternateRecipientAllowed", 0x0002);
        /// <summary>PidTagImportance (0x0017).</summary>
        public static readonly MapiPropertyKey<int> Importance = Integer("PidTagImportance", 0x0017);
        /// <summary>PidTagMessageClass (0x001A).</summary>
        public static readonly MapiPropertyKey<string> MessageClass = String("PidTagMessageClass", 0x001A);
        /// <summary>PidTagOriginatorDeliveryReportRequested (0x0023).</summary>
        public static readonly MapiPropertyKey<bool> OriginatorDeliveryReportRequested =
            Boolean("PidTagOriginatorDeliveryReportRequested", 0x0023);
        /// <summary>PidTagPriority (0x0026).</summary>
        public static readonly MapiPropertyKey<int> Priority = Integer("PidTagPriority", 0x0026);
        /// <summary>PidTagReadReceiptRequested (0x0029).</summary>
        public static readonly MapiPropertyKey<bool> ReadReceiptRequested = Boolean("PidTagReadReceiptRequested", 0x0029);
        /// <summary>PidTagOriginalSensitivity (0x002E).</summary>
        public static readonly MapiPropertyKey<int> OriginalSensitivity = Integer("PidTagOriginalSensitivity", 0x002E);
        /// <summary>PidTagSensitivity (0x0036).</summary>
        public static readonly MapiPropertyKey<int> Sensitivity = Integer("PidTagSensitivity", 0x0036);
        /// <summary>PidTagSubject (0x0037).</summary>
        public static readonly MapiPropertyKey<string> Subject = String("PidTagSubject", 0x0037);
        /// <summary>PidTagClientSubmitTime (0x0039).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> ClientSubmitTime = Time("PidTagClientSubmitTime", 0x0039);
        /// <summary>PidTagSubjectPrefix (0x003D).</summary>
        public static readonly MapiPropertyKey<string> SubjectPrefix = String("PidTagSubjectPrefix", 0x003D);
        /// <summary>PidTagSentRepresentingName (0x0042).</summary>
        public static readonly MapiPropertyKey<string> SentRepresentingName = String("PidTagSentRepresentingName", 0x0042);
        /// <summary>PidTagReplyRecipientNames (0x0050).</summary>
        public static readonly MapiPropertyKey<string> ReplyRecipientNames = String("PidTagReplyRecipientNames", 0x0050);
        /// <summary>PidTagSentRepresentingAddressType (0x0064).</summary>
        public static readonly MapiPropertyKey<string> SentRepresentingAddressType =
            String("PidTagSentRepresentingAddressType", 0x0064);
        /// <summary>PidTagSentRepresentingEmailAddress (0x0065).</summary>
        public static readonly MapiPropertyKey<string> SentRepresentingEmailAddress =
            String("PidTagSentRepresentingEmailAddress", 0x0065);
        /// <summary>PidTagConversationTopic (0x0070).</summary>
        public static readonly MapiPropertyKey<string> ConversationTopic = String("PidTagConversationTopic", 0x0070);
        /// <summary>PidTagConversationIndex (0x0071).</summary>
        public static readonly MapiPropertyKey<byte[]> ConversationIndex = Binary("PidTagConversationIndex", 0x0071);
        /// <summary>PidTagTransportMessageHeaders (0x007D).</summary>
        public static readonly MapiPropertyKey<string> TransportMessageHeaders =
            String("PidTagTransportMessageHeaders", 0x007D);
        /// <summary>PidTagSenderEntryId (0x0C19).</summary>
        public static readonly MapiPropertyKey<byte[]> SenderEntryId = Binary("PidTagSenderEntryId", 0x0C19);
        /// <summary>PidTagSenderName (0x0C1A).</summary>
        public static readonly MapiPropertyKey<string> SenderName = String("PidTagSenderName", 0x0C1A);
        /// <summary>PidTagRecipientType (0x0C15).</summary>
        public static readonly MapiPropertyKey<int> RecipientType = Integer("PidTagRecipientType", 0x0C15);
        /// <summary>PidTagSenderAddressType (0x0C1E).</summary>
        public static readonly MapiPropertyKey<string> SenderAddressType = String("PidTagSenderAddressType", 0x0C1E);
        /// <summary>PidTagSenderEmailAddress (0x0C1F).</summary>
        public static readonly MapiPropertyKey<string> SenderEmailAddress = String("PidTagSenderEmailAddress", 0x0C1F);
        /// <summary>PidTagDisplayBcc (0x0E02).</summary>
        public static readonly MapiPropertyKey<string> DisplayBcc = String("PidTagDisplayBcc", 0x0E02);
        /// <summary>PidTagDisplayCc (0x0E03).</summary>
        public static readonly MapiPropertyKey<string> DisplayCc = String("PidTagDisplayCc", 0x0E03);
        /// <summary>PidTagDisplayTo (0x0E04).</summary>
        public static readonly MapiPropertyKey<string> DisplayTo = String("PidTagDisplayTo", 0x0E04);
        /// <summary>PidTagMessageDeliveryTime (0x0E06).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> MessageDeliveryTime = Time("PidTagMessageDeliveryTime", 0x0E06);
        /// <summary>PidTagMessageFlags (0x0E07).</summary>
        public static readonly MapiPropertyKey<int> MessageFlags = Integer("PidTagMessageFlags", 0x0E07);
        /// <summary>PidTagMessageSize (0x0E08).</summary>
        public static readonly MapiPropertyKey<int> MessageSize = Integer("PidTagMessageSize", 0x0E08);
        /// <summary>PidTagResponsibility (0x0E0F).</summary>
        public static readonly MapiPropertyKey<bool> Responsibility = Boolean("PidTagResponsibility", 0x0E0F);
        /// <summary>PidTagMessageStatus (0x0E17).</summary>
        public static readonly MapiPropertyKey<int> MessageStatus = Integer("PidTagMessageStatus", 0x0E17);
        /// <summary>PidTagHasAttachments (0x0E1B).</summary>
        public static readonly MapiPropertyKey<bool> HasAttachments = Boolean("PidTagHasAttachments", 0x0E1B);
        /// <summary>PidTagNormalizedSubject (0x0E1D).</summary>
        public static readonly MapiPropertyKey<string> NormalizedSubject = String("PidTagNormalizedSubject", 0x0E1D);
        /// <summary>PidTagRtfInSync (0x0E1F).</summary>
        public static readonly MapiPropertyKey<bool> RtfInSync = Boolean("PidTagRtfInSync", 0x0E1F);
        /// <summary>PidTagAttachSize (0x0E20).</summary>
        public static readonly MapiPropertyKey<int> AttachSize = Integer("PidTagAttachSize", 0x0E20);
        /// <summary>PidTagAttachNumber (0x0E21).</summary>
        public static readonly MapiPropertyKey<int> AttachNumber = Integer("PidTagAttachNumber", 0x0E21);
        /// <summary>PidTagObjectType (0x0FFE).</summary>
        public static readonly MapiPropertyKey<int> ObjectType = Integer("PidTagObjectType", 0x0FFE);
        /// <summary>PidTagEntryId (0x0FFF).</summary>
        public static readonly MapiPropertyKey<byte[]> EntryId = Binary("PidTagEntryId", 0x0FFF);
        /// <summary>PidTagRecordKey (0x0FF9).</summary>
        public static readonly MapiPropertyKey<byte[]> RecordKey = Binary("PidTagRecordKey", 0x0FF9);
        /// <summary>PidTagBody (0x1000).</summary>
        public static readonly MapiPropertyKey<string> Body = String("PidTagBody", 0x1000);
        /// <summary>PidTagRtfCompressed (0x1009).</summary>
        public static readonly MapiPropertyKey<byte[]> RtfCompressed = Binary("PidTagRtfCompressed", 0x1009);
        /// <summary>PidTagHtml (0x1013), which can be retained as binary or decoded text.</summary>
        public static readonly MapiPropertyKey<object> Html = new MapiPropertyKey<object>("PidTagHtml", 0x1013,
            MapiPropertyType.Binary, MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidTagNativeBodyInfo (0x1016).</summary>
        public static readonly MapiPropertyKey<int> NativeBodyInfo = Integer("PidTagNativeBodyInfo", 0x1016);
        /// <summary>PidTagInternetMessageId (0x1035).</summary>
        public static readonly MapiPropertyKey<string> InternetMessageId = String("PidTagInternetMessageId", 0x1035);
        /// <summary>PidTagInternetReferences (0x1039).</summary>
        public static readonly MapiPropertyKey<string> InternetReferences = String("PidTagInternetReferences", 0x1039);
        /// <summary>PidTagInReplyToId (0x1042).</summary>
        public static readonly MapiPropertyKey<string> InReplyToId = String("PidTagInReplyToId", 0x1042);
        /// <summary>PidTagIconIndex (0x1080).</summary>
        public static readonly MapiPropertyKey<int> IconIndex = Integer("PidTagIconIndex", 0x1080);
        /// <summary>PidTagItemTemporaryFlags (0x1097).</summary>
        public static readonly MapiPropertyKey<int> ItemTemporaryFlags = Integer("PidTagItemTemporaryFlags", 0x1097);
        /// <summary>PidTagRowId (0x3000).</summary>
        public static readonly MapiPropertyKey<int> RowId = Integer("PidTagRowId", 0x3000);
        /// <summary>PidTagDisplayName (0x3001).</summary>
        public static readonly MapiPropertyKey<string> DisplayName = String("PidTagDisplayName", 0x3001);
        /// <summary>PidTagAddressType (0x3002).</summary>
        public static readonly MapiPropertyKey<string> AddressType = String("PidTagAddressType", 0x3002);
        /// <summary>PidTagEmailAddress (0x3003).</summary>
        public static readonly MapiPropertyKey<string> EmailAddress = String("PidTagEmailAddress", 0x3003);
        /// <summary>PidTagCreationTime (0x3007).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> CreationTime = Time("PidTagCreationTime", 0x3007);
        /// <summary>PidTagLastModificationTime (0x3008).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> LastModificationTime = Time("PidTagLastModificationTime", 0x3008);
        /// <summary>PidTagSearchKey (0x300B).</summary>
        public static readonly MapiPropertyKey<byte[]> SearchKey = Binary("PidTagSearchKey", 0x300B);
        /// <summary>PidTagConversationId (0x3013).</summary>
        public static readonly MapiPropertyKey<byte[]> ConversationId = Binary("PidTagConversationId", 0x3013);
        /// <summary>PidTagStoreSupportMask (0x340D).</summary>
        public static readonly MapiPropertyKey<int> StoreSupportMask = Integer("PidTagStoreSupportMask", 0x340D);
        /// <summary>PidTagContentCount (0x3602).</summary>
        public static readonly MapiPropertyKey<int> ContentCount = Integer("PidTagContentCount", 0x3602);
        /// <summary>PidTagContentUnreadCount (0x3603).</summary>
        public static readonly MapiPropertyKey<int> ContentUnreadCount = Integer("PidTagContentUnreadCount", 0x3603);
        /// <summary>PidTagSubfolders (0x360A).</summary>
        public static readonly MapiPropertyKey<bool> Subfolders = Boolean("PidTagSubfolders", 0x360A);
        /// <summary>PidTagContainerClass (0x3613).</summary>
        public static readonly MapiPropertyKey<string> ContainerClass = String("PidTagContainerClass", 0x3613);
        /// <summary>PidTagAssociatedContentCount (0x3617).</summary>
        public static readonly MapiPropertyKey<int> AssociatedContentCount = Integer("PidTagAssociatedContentCount", 0x3617);
        /// <summary>PidTagAttachData (0x3701).</summary>
        public static readonly MapiPropertyKey<object> AttachData = new MapiPropertyKey<object>("PidTagAttachData", 0x3701,
            MapiPropertyType.Binary, MapiPropertyType.Object);
        /// <summary>PidTagAttachExtension (0x3703).</summary>
        public static readonly MapiPropertyKey<string> AttachExtension = String("PidTagAttachExtension", 0x3703);
        /// <summary>PidTagAttachFilename (0x3704).</summary>
        public static readonly MapiPropertyKey<string> AttachFilename = String("PidTagAttachFilename", 0x3704);
        /// <summary>PidTagAttachMethod (0x3705).</summary>
        public static readonly MapiPropertyKey<int> AttachMethod = Integer("PidTagAttachMethod", 0x3705);
        /// <summary>PidTagAttachLongFilename (0x3707).</summary>
        public static readonly MapiPropertyKey<string> AttachLongFilename = String("PidTagAttachLongFilename", 0x3707);
        /// <summary>PidTagRenderingPosition (0x370B).</summary>
        public static readonly MapiPropertyKey<int> RenderingPosition = Integer("PidTagRenderingPosition", 0x370B);
        /// <summary>PidTagAttachLongPathname (0x370D).</summary>
        public static readonly MapiPropertyKey<string> AttachLongPathname = String("PidTagAttachLongPathname", 0x370D);
        /// <summary>PidTagAttachMimeTag (0x370E).</summary>
        public static readonly MapiPropertyKey<string> AttachMimeTag = String("PidTagAttachMimeTag", 0x370E);
        /// <summary>PidTagAttachContentId (0x3712).</summary>
        public static readonly MapiPropertyKey<string> AttachContentId = String("PidTagAttachContentId", 0x3712);
        /// <summary>PidTagAttachContentLocation (0x3713).</summary>
        public static readonly MapiPropertyKey<string> AttachContentLocation = String("PidTagAttachContentLocation", 0x3713);
        /// <summary>PidTagAttachFlags (0x3714).</summary>
        public static readonly MapiPropertyKey<int> AttachFlags = Integer("PidTagAttachFlags", 0x3714);
        /// <summary>PidTagDisplayType (0x3900).</summary>
        public static readonly MapiPropertyKey<int> DisplayType = Integer("PidTagDisplayType", 0x3900);
        /// <summary>PidTagDisplayTypeEx (0x3905).</summary>
        public static readonly MapiPropertyKey<int> DisplayTypeEx = Integer("PidTagDisplayTypeEx", 0x3905);
        /// <summary>PidTagSmtpAddress (0x39FE).</summary>
        public static readonly MapiPropertyKey<string> SmtpAddress = String("PidTagSmtpAddress", 0x39FE);
        /// <summary>PidTagDisplayNamePrintable (0x39FF).</summary>
        public static readonly MapiPropertyKey<string> DisplayNamePrintable = String("PidTagDisplayNamePrintable", 0x39FF);
        /// <summary>PidTagSendRichInfo (0x3A40).</summary>
        public static readonly MapiPropertyKey<bool> SendRichInfo = Boolean("PidTagSendRichInfo", 0x3A40);
        /// <summary>PidTagInternetCodepage (0x3FDE).</summary>
        public static readonly MapiPropertyKey<int> InternetCodepage = Integer("PidTagInternetCodepage", 0x3FDE);
        /// <summary>PidTagMessageLocaleId (0x3FF1).</summary>
        public static readonly MapiPropertyKey<int> MessageLocaleId = Integer("PidTagMessageLocaleId", 0x3FF1);
        /// <summary>PidTagCodePageId (0x3FFC).</summary>
        public static readonly MapiPropertyKey<int> CodePageId = Integer("PidTagCodePageId", 0x3FFC);
        /// <summary>PidTagMessageCodepage (0x3FFD).</summary>
        public static readonly MapiPropertyKey<int> MessageCodepage = Integer("PidTagMessageCodepage", 0x3FFD);
        /// <summary>PidTagLastModifierName (0x3FFA).</summary>
        public static readonly MapiPropertyKey<string> LastModifierName = String("PidTagLastModifierName", 0x3FFA);
        /// <summary>PidTagMessageEditorFormat (0x5909).</summary>
        public static readonly MapiPropertyKey<int> MessageEditorFormat = Integer("PidTagMessageEditorFormat", 0x5909);
        /// <summary>PidTagSenderSmtpAddress (0x5D01).</summary>
        public static readonly MapiPropertyKey<string> SenderSmtpAddress = String("PidTagSenderSmtpAddress", 0x5D01);
        /// <summary>PidTagSentRepresentingSmtpAddress (0x5D02).</summary>
        public static readonly MapiPropertyKey<string> SentRepresentingSmtpAddress =
            String("PidTagSentRepresentingSmtpAddress", 0x5D02);
        /// <summary>PidTagRecipientDisplayName (0x5FF6).</summary>
        public static readonly MapiPropertyKey<string> RecipientDisplayName =
            String("PidTagRecipientDisplayName", 0x5FF6);
        /// <summary>PidTagLtpRowId (0x67F2).</summary>
        public static readonly MapiPropertyKey<int> LtpRowId = Integer("PidTagLtpRowId", 0x67F2);
        /// <summary>PidTagLtpRowVer (0x67F3).</summary>
        public static readonly MapiPropertyKey<int> LtpRowVer = Integer("PidTagLtpRowVer", 0x67F3);
        /// <summary>PidTagPstPassword (0x67FF).</summary>
        public static readonly MapiPropertyKey<int> PstPassword = Integer("PidTagPstPassword", 0x67FF);
        /// <summary>PidTagAttachmentHidden (0x7FFE).</summary>
        public static readonly MapiPropertyKey<bool> AttachmentHidden = Boolean("PidTagAttachmentHidden", 0x7FFE);
        /// <summary>PidTagAttachmentContactPhoto (0x7FFF).</summary>
        public static readonly MapiPropertyKey<bool> AttachmentContactPhoto = Boolean("PidTagAttachmentContactPhoto", 0x7FFF);

        private static readonly IReadOnlyList<MapiPropertyKey> CoreProperties = new MapiPropertyKey[] {
            NameidBucketCount, NameidStreamGuid, NameidStreamEntry, NameidStreamString, AlternateRecipientAllowed,
            Importance, MessageClass, OriginatorDeliveryReportRequested, Priority, ReadReceiptRequested,
            OriginalSensitivity, Sensitivity, Subject, ClientSubmitTime, SubjectPrefix, SentRepresentingName,
            ReplyRecipientNames, SentRepresentingAddressType, SentRepresentingEmailAddress, ConversationTopic,
            ConversationIndex, TransportMessageHeaders, SenderEntryId, SenderName, RecipientType, SenderAddressType,
            SenderEmailAddress, DisplayBcc, DisplayCc, DisplayTo, MessageDeliveryTime, MessageFlags,
            MessageSize, Responsibility, MessageStatus, HasAttachments, NormalizedSubject, RtfInSync, AttachSize,
            AttachNumber, ObjectType, EntryId, RecordKey, Body, RtfCompressed, Html, NativeBodyInfo,
            InternetMessageId, InternetReferences, InReplyToId, IconIndex, ItemTemporaryFlags, RowId, DisplayName,
            AddressType, EmailAddress, CreationTime, LastModificationTime, SearchKey, ConversationId, StoreSupportMask,
            ContentCount, ContentUnreadCount, Subfolders,
            ContainerClass, AssociatedContentCount, AttachData, AttachExtension, AttachFilename, AttachMethod,
            AttachLongFilename, RenderingPosition, AttachLongPathname, AttachMimeTag, AttachContentId,
            AttachContentLocation, AttachFlags, DisplayType, DisplayTypeEx, SmtpAddress, DisplayNamePrintable,
            SendRichInfo, InternetCodepage, MessageLocaleId, CodePageId, MessageCodepage, LastModifierName,
            MessageEditorFormat, SenderSmtpAddress, SentRepresentingSmtpAddress, RecipientDisplayName, LtpRowId,
            LtpRowVer, PstPassword, AttachmentHidden, AttachmentContactPhoto
        };

        internal static IReadOnlyList<MapiPropertyKey> All => CoreProperties
            .Concat(PstAndTableProperties)
            .Concat(OutlookTaggedProperties)
            .Concat(OabTaggedProperties)
            .Concat(InfrastructureTaggedProperties)
            .ToArray();
    }

    /// <summary>Well-known numeric named properties (PidLid).</summary>
    public static partial class PidLid {
        /// <summary>PidLidReminderDelta (0x8501 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<int> ReminderDelta = new MapiPropertyKey<int>("PidLidReminderDelta",
            MapiPropertySets.Common, 0x8501, MapiPropertyType.Integer32);
        /// <summary>PidLidReminderTime (0x8502 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> ReminderTime = new MapiPropertyKey<DateTimeOffset>(
            "PidLidReminderTime", MapiPropertySets.Common, 0x8502, MapiPropertyType.Time);
        /// <summary>PidLidReminderSet (0x8503 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<bool> ReminderSet = new MapiPropertyKey<bool>("PidLidReminderSet",
            MapiPropertySets.Common, 0x8503, MapiPropertyType.Boolean);
        /// <summary>PidLidPrivate (0x8506 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<bool> Private = new MapiPropertyKey<bool>("PidLidPrivate",
            MapiPropertySets.Common, 0x8506, MapiPropertyType.Boolean);
        /// <summary>PidLidCommonStart (0x8516 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> CommonStart = new MapiPropertyKey<DateTimeOffset>(
            "PidLidCommonStart", MapiPropertySets.Common, 0x8516, MapiPropertyType.Time);
        /// <summary>PidLidCommonEnd (0x8517 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> CommonEnd = new MapiPropertyKey<DateTimeOffset>(
            "PidLidCommonEnd", MapiPropertySets.Common, 0x8517, MapiPropertyType.Time);
        /// <summary>PidLidReminderOverride (0x851C in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<bool> ReminderOverride = new MapiPropertyKey<bool>(
            "PidLidReminderOverride", MapiPropertySets.Common, 0x851C, MapiPropertyType.Boolean);
        /// <summary>PidLidReminderPlaySound (0x851E in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<bool> ReminderPlaySound = new MapiPropertyKey<bool>(
            "PidLidReminderPlaySound", MapiPropertySets.Common, 0x851E, MapiPropertyType.Boolean);
        /// <summary>PidLidReminderFileParameter (0x851F in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<string> ReminderFileParameter = new MapiPropertyKey<string>(
            "PidLidReminderFileParameter", MapiPropertySets.Common, 0x851F,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidVerbStream (0x8520 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<byte[]> VerbStream = new MapiPropertyKey<byte[]>(
            "PidLidVerbStream", MapiPropertySets.Common, 0x8520, MapiPropertyType.Binary);
        /// <summary>PidLidVerbResponse (0x8524 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<string> VerbResponse = new MapiPropertyKey<string>(
            "PidLidVerbResponse", MapiPropertySets.Common, 0x8524,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidFlagRequest (0x8530 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<string> FlagRequest = new MapiPropertyKey<string>(
            "PidLidFlagRequest", MapiPropertySets.Common, 0x8530,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidPropertyDefinitionStream (0x8540 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<byte[]> PropertyDefinitionStream = new MapiPropertyKey<byte[]>(
            "PidLidPropertyDefinitionStream", MapiPropertySets.Common, 0x8540, MapiPropertyType.Binary);
        /// <summary>PidLidReminderSignalTime (0x8560 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> ReminderSignalTime = new MapiPropertyKey<DateTimeOffset>(
            "PidLidReminderSignalTime", MapiPropertySets.Common, 0x8560, MapiPropertyType.Time);
        /// <summary>PidLidToDoTitle (0x85A4 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<string> ToDoTitle = new MapiPropertyKey<string>(
            "PidLidToDoTitle", MapiPropertySets.Common, 0x85A4,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidValidFlagStringProof (0x85BF in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> ValidFlagStringProof = new MapiPropertyKey<DateTimeOffset>(
            "PidLidValidFlagStringProof", MapiPropertySets.Common, 0x85BF, MapiPropertyType.Time);
        /// <summary>PidLidFlagString (0x85C0 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<int> FlagString = new MapiPropertyKey<int>(
            "PidLidFlagString", MapiPropertySets.Common, 0x85C0, MapiPropertyType.Integer32);
        /// <summary>PidLidSideEffects (0x8510 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<int> SideEffects = new MapiPropertyKey<int>("PidLidSideEffects",
            MapiPropertySets.Common, 0x8510, MapiPropertyType.Integer32);
        /// <summary>PidLidHeaderItem (0x8578 in PSETID_Common).</summary>
        public static readonly MapiPropertyKey<bool> HeaderItem = new MapiPropertyKey<bool>("PidLidHeaderItem",
            MapiPropertySets.Common, 0x8578, MapiPropertyType.Boolean,
            MapiPropertyType.Integer32, MapiPropertyType.Integer16);
        /// <summary>PidLidAppointmentStartWhole (0x820D in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> AppointmentStartWhole = new MapiPropertyKey<DateTimeOffset>(
            "PidLidAppointmentStartWhole", MapiPropertySets.Appointment, 0x820D, MapiPropertyType.Time);
        /// <summary>PidLidAppointmentEndWhole (0x820E in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> AppointmentEndWhole = new MapiPropertyKey<DateTimeOffset>(
            "PidLidAppointmentEndWhole", MapiPropertySets.Appointment, 0x820E, MapiPropertyType.Time);
        /// <summary>PidLidLocation (0x8208 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<string> Location = new MapiPropertyKey<string>("PidLidLocation",
            MapiPropertySets.Appointment, 0x8208, MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidAppointmentSubType (0x8215 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<bool> AppointmentSubType = new MapiPropertyKey<bool>(
            "PidLidAppointmentSubType", MapiPropertySets.Appointment, 0x8215, MapiPropertyType.Boolean);
        /// <summary>PidLidBusyStatus (0x8205 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<int> BusyStatus = new MapiPropertyKey<int>("PidLidBusyStatus",
            MapiPropertySets.Appointment, 0x8205, MapiPropertyType.Integer32);
        /// <summary>PidLidAppointmentSequence (0x8201 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<int> AppointmentSequence = new MapiPropertyKey<int>(
            "PidLidAppointmentSequence", MapiPropertySets.Appointment, 0x8201, MapiPropertyType.Integer32);
        /// <summary>PidLidAppointmentDuration (0x8213 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<int> AppointmentDuration = new MapiPropertyKey<int>(
            "PidLidAppointmentDuration", MapiPropertySets.Appointment, 0x8213, MapiPropertyType.Integer32);
        /// <summary>PidLidAppointmentRecur (0x8216 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<byte[]> AppointmentRecur = new MapiPropertyKey<byte[]>(
            "PidLidAppointmentRecur", MapiPropertySets.Appointment, 0x8216, MapiPropertyType.Binary);
        /// <summary>PidLidAppointmentStateFlags (0x8217 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<int> AppointmentStateFlags = new MapiPropertyKey<int>(
            "PidLidAppointmentStateFlags", MapiPropertySets.Appointment, 0x8217, MapiPropertyType.Integer32);
        /// <summary>PidLidResponseStatus (0x8218 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<int> ResponseStatus = new MapiPropertyKey<int>(
            "PidLidResponseStatus", MapiPropertySets.Appointment, 0x8218, MapiPropertyType.Integer32);
        /// <summary>PidLidRecurring (0x8223 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<bool> Recurring = new MapiPropertyKey<bool>("PidLidRecurring",
            MapiPropertySets.Appointment, 0x8223, MapiPropertyType.Boolean);
        /// <summary>PidLidAppointmentReplyTime (0x8220 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> AppointmentReplyTime = new MapiPropertyKey<DateTimeOffset>(
            "PidLidAppointmentReplyTime", MapiPropertySets.Appointment, 0x8220, MapiPropertyType.Time);
        /// <summary>PidLidIntendedBusyStatus (0x8224 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<int> IntendedBusyStatus = new MapiPropertyKey<int>(
            "PidLidIntendedBusyStatus", MapiPropertySets.Appointment, 0x8224, MapiPropertyType.Integer32);
        /// <summary>PidLidAppointmentReplyName (0x8230 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<string> AppointmentReplyName = new MapiPropertyKey<string>(
            "PidLidAppointmentReplyName", MapiPropertySets.Appointment, 0x8230,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidRecurrenceType (0x8231 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<int> RecurrenceType = new MapiPropertyKey<int>("PidLidRecurrenceType",
            MapiPropertySets.Appointment, 0x8231, MapiPropertyType.Integer32);
        /// <summary>PidLidRecurrencePattern (0x8232 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<string> RecurrencePattern = new MapiPropertyKey<string>(
            "PidLidRecurrencePattern", MapiPropertySets.Appointment, 0x8232,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidTimeZoneStruct (0x8233 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<byte[]> TimeZoneStruct = new MapiPropertyKey<byte[]>("PidLidTimeZoneStruct",
            MapiPropertySets.Appointment, 0x8233, MapiPropertyType.Binary);
        /// <summary>PidLidTimeZoneDescription (0x8234 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<string> TimeZoneDescription = new MapiPropertyKey<string>(
            "PidLidTimeZoneDescription", MapiPropertySets.Appointment, 0x8234,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidAllAttendeesString (0x8238 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<string> AllAttendeesString = new MapiPropertyKey<string>(
            "PidLidAllAttendeesString", MapiPropertySets.Appointment, 0x8238,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidToAttendeesString (0x823B in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<string> ToAttendeesString = new MapiPropertyKey<string>(
            "PidLidToAttendeesString", MapiPropertySets.Appointment, 0x823B,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidCcAttendeesString (0x823C in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<string> CcAttendeesString = new MapiPropertyKey<string>(
            "PidLidCcAttendeesString", MapiPropertySets.Appointment, 0x823C,
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidAppointmentNotAllowPropose (0x825A in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<bool> AppointmentNotAllowPropose = new MapiPropertyKey<bool>(
            "PidLidAppointmentNotAllowPropose", MapiPropertySets.Appointment, 0x825A, MapiPropertyType.Boolean);
        /// <summary>PidLidAppointmentProposedStartWhole (0x8250 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> AppointmentProposedStartWhole =
            new MapiPropertyKey<DateTimeOffset>("PidLidAppointmentProposedStartWhole",
                MapiPropertySets.Appointment, 0x8250, MapiPropertyType.Time);
        /// <summary>PidLidAppointmentProposedEndWhole (0x8251 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> AppointmentProposedEndWhole =
            new MapiPropertyKey<DateTimeOffset>("PidLidAppointmentProposedEndWhole",
                MapiPropertySets.Appointment, 0x8251, MapiPropertyType.Time);
        /// <summary>PidLidAppointmentProposedDuration (0x8256 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<int> AppointmentProposedDuration = new MapiPropertyKey<int>(
            "PidLidAppointmentProposedDuration", MapiPropertySets.Appointment, 0x8256, MapiPropertyType.Integer32);
        /// <summary>PidLidAppointmentCounterProposal (0x8257 in PSETID_Appointment).</summary>
        public static readonly MapiPropertyKey<bool> AppointmentCounterProposal = new MapiPropertyKey<bool>(
            "PidLidAppointmentCounterProposal", MapiPropertySets.Appointment, 0x8257, MapiPropertyType.Boolean);
        /// <summary>PidLidAppointmentTimeZoneDefinitionStartDisplay (0x825E).</summary>
        public static readonly MapiPropertyKey<byte[]> AppointmentTimeZoneDefinitionStartDisplay =
            new MapiPropertyKey<byte[]>("PidLidAppointmentTimeZoneDefinitionStartDisplay",
                MapiPropertySets.Appointment, 0x825E, MapiPropertyType.Binary);
        /// <summary>PidLidAppointmentTimeZoneDefinitionEndDisplay (0x825F).</summary>
        public static readonly MapiPropertyKey<byte[]> AppointmentTimeZoneDefinitionEndDisplay =
            new MapiPropertyKey<byte[]>("PidLidAppointmentTimeZoneDefinitionEndDisplay",
                MapiPropertySets.Appointment, 0x825F, MapiPropertyType.Binary);
        /// <summary>PidLidAppointmentTimeZoneDefinitionRecur (0x8260).</summary>
        public static readonly MapiPropertyKey<byte[]> AppointmentTimeZoneDefinitionRecur =
            new MapiPropertyKey<byte[]>("PidLidAppointmentTimeZoneDefinitionRecur",
                MapiPropertySets.Appointment, 0x8260, MapiPropertyType.Binary);
        /// <summary>PidLidClientIntent (0x0015 in the calendar-assistant property set).</summary>
        public static readonly MapiPropertyKey<int> ClientIntent = new MapiPropertyKey<int>("PidLidClientIntent",
            MapiPropertySets.CalendarAssistant, 0x0015, MapiPropertyType.Integer32);
        /// <summary>PidLidAttendeeCriticalChange (0x0001 in PSETID_Meeting).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> AttendeeCriticalChange =
            new MapiPropertyKey<DateTimeOffset>("PidLidAttendeeCriticalChange",
                MapiPropertySets.Meeting, 0x0001, MapiPropertyType.Time);
        /// <summary>PidLidGlobalObjectId (0x0003 in PSETID_Meeting).</summary>
        public static readonly MapiPropertyKey<byte[]> GlobalObjectId = new MapiPropertyKey<byte[]>(
            "PidLidGlobalObjectId", MapiPropertySets.Meeting, 0x0003, MapiPropertyType.Binary);
        /// <summary>PidLidIsSilent (0x0004 in PSETID_Meeting).</summary>
        public static readonly MapiPropertyKey<bool> IsSilent = new MapiPropertyKey<bool>("PidLidIsSilent",
            MapiPropertySets.Meeting, 0x0004, MapiPropertyType.Boolean);
        /// <summary>PidLidOwnerCriticalChange (0x001A in PSETID_Meeting).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> OwnerCriticalChange =
            new MapiPropertyKey<DateTimeOffset>("PidLidOwnerCriticalChange",
                MapiPropertySets.Meeting, 0x001A, MapiPropertyType.Time);
        /// <summary>PidLidCleanGlobalObjectId (0x0023 in PSETID_Meeting).</summary>
        public static readonly MapiPropertyKey<byte[]> CleanGlobalObjectId = new MapiPropertyKey<byte[]>(
            "PidLidCleanGlobalObjectId", MapiPropertySets.Meeting, 0x0023, MapiPropertyType.Binary);
        /// <summary>PidLidMeetingType (0x0026 in PSETID_Meeting).</summary>
        public static readonly MapiPropertyKey<int> MeetingType = new MapiPropertyKey<int>("PidLidMeetingType",
            MapiPropertySets.Meeting, 0x0026, MapiPropertyType.Integer32);

        private static readonly IReadOnlyList<MapiPropertyKey> CoreProperties = new MapiPropertyKey[] {
            ReminderDelta, ReminderTime, ReminderSet, Private, CommonStart, CommonEnd, ReminderOverride,
            ReminderPlaySound, ReminderFileParameter, VerbStream, VerbResponse, FlagRequest, ReminderSignalTime,
            ToDoTitle, ValidFlagStringProof, FlagString, PropertyDefinitionStream, SideEffects,
            HeaderItem,
            AppointmentStartWhole, AppointmentEndWhole, Location, AppointmentSubType, BusyStatus,
            AppointmentSequence, AppointmentDuration, AppointmentRecur, AppointmentStateFlags, ResponseStatus,
            AppointmentReplyTime, IntendedBusyStatus, AppointmentReplyName,
            Recurring, RecurrenceType, RecurrencePattern, TimeZoneStruct, TimeZoneDescription, AllAttendeesString,
            ToAttendeesString, CcAttendeesString, AppointmentNotAllowPropose, AppointmentProposedStartWhole,
            AppointmentProposedEndWhole, AppointmentProposedDuration, AppointmentCounterProposal,
            AppointmentTimeZoneDefinitionStartDisplay, AppointmentTimeZoneDefinitionEndDisplay,
            AppointmentTimeZoneDefinitionRecur, ClientIntent, AttendeeCriticalChange, GlobalObjectId, IsSilent,
            OwnerCriticalChange, CleanGlobalObjectId, MeetingType
        };

        internal static IReadOnlyList<MapiPropertyKey> All => CoreProperties
            .Concat(OutlookNamedProperties)
            .ToArray();
    }

    /// <summary>Well-known string named properties (PidName).</summary>
    public static class PidName {
        /// <summary>PidNameKeywords in PS_PUBLIC_STRINGS.</summary>
        public static readonly MapiPropertyKey<object> Keywords = new MapiPropertyKey<object>("PidNameKeywords",
            MapiPropertySets.PublicStrings, "Keywords", MapiPropertyType.MultipleUnicode,
            MapiPropertyType.MultipleString8, MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidNameAcceptLanguage in PS_INTERNET_HEADERS.</summary>
        public static readonly MapiPropertyKey<string> AcceptLanguage = new MapiPropertyKey<string>(
            "PidNameAcceptLanguage", MapiPropertySets.InternetHeaders, "acceptlanguage",
            MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>Outlook ReactionsSummary property.</summary>
        public static readonly MapiPropertyKey<byte[]> ReactionsSummary = new MapiPropertyKey<byte[]>(
            "ReactionsSummary", MapiPropertySets.Reactions, "ReactionsSummary", MapiPropertyType.Binary);
        /// <summary>Outlook OwnerReactionHistory property.</summary>
        public static readonly MapiPropertyKey<byte[]> OwnerReactionHistory = new MapiPropertyKey<byte[]>(
            "OwnerReactionHistory", MapiPropertySets.Reactions, "OwnerReactionHistory", MapiPropertyType.Binary);
        /// <summary>Outlook OwnerReactionType property.</summary>
        public static readonly MapiPropertyKey<string> OwnerReactionType = new MapiPropertyKey<string>(
            "OwnerReactionType", MapiPropertySets.Reactions, "OwnerReactionType", MapiPropertyType.Unicode,
            MapiPropertyType.String8);
        /// <summary>Outlook OwnerReactionTime property.</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> OwnerReactionTime = new MapiPropertyKey<DateTimeOffset>(
            "OwnerReactionTime", MapiPropertySets.Reactions, "OwnerReactionTime", MapiPropertyType.Time);
        /// <summary>Outlook ReactionsCount property.</summary>
        public static readonly MapiPropertyKey<int> ReactionsCount = new MapiPropertyKey<int>("ReactionsCount",
            MapiPropertySets.Reactions, "ReactionsCount", MapiPropertyType.Integer32);
        /// <summary>OfficeIMO PST writer provenance marker.</summary>
        public static readonly MapiPropertyKey<string> OfficeImoPstWriter = new MapiPropertyKey<string>(
            "OfficeImoPstWriter", MapiPropertySets.OfficeImoEmailStore, "OfficeIMO.PstWriter",
            MapiPropertyType.Unicode, MapiPropertyType.String8);

        internal static readonly IReadOnlyList<MapiPropertyKey> All = new MapiPropertyKey[] {
            Keywords, AcceptLanguage, ReactionsSummary, OwnerReactionHistory, OwnerReactionType, OwnerReactionTime,
            ReactionsCount, OfficeImoPstWriter
        };
    }
}
