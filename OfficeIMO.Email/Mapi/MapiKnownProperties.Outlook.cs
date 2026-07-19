namespace OfficeIMO.Email;

public static partial class MapiKnownProperties {
    public static partial class PidTag {
        /// <summary>PidTagReceivedByEntryId (0x003F).</summary>
        public static readonly MapiPropertyKey<byte[]> ReceivedByEntryId = Binary("PidTagReceivedByEntryId", 0x003F);
        /// <summary>PidTagReceivedByName (0x0040).</summary>
        public static readonly MapiPropertyKey<string> ReceivedByName = String("PidTagReceivedByName", 0x0040);
        /// <summary>PidTagReceivedRepresentingEntryId (0x0043).</summary>
        public static readonly MapiPropertyKey<byte[]> ReceivedRepresentingEntryId =
            Binary("PidTagReceivedRepresentingEntryId", 0x0043);
        /// <summary>PidTagReceivedRepresentingName (0x0044).</summary>
        public static readonly MapiPropertyKey<string> ReceivedRepresentingName =
            String("PidTagReceivedRepresentingName", 0x0044);
        /// <summary>PidTagReceivedByAddressType (0x0075).</summary>
        public static readonly MapiPropertyKey<string> ReceivedByAddressType =
            String("PidTagReceivedByAddressType", 0x0075);
        /// <summary>PidTagReceivedByEmailAddress (0x0076).</summary>
        public static readonly MapiPropertyKey<string> ReceivedByEmailAddress =
            String("PidTagReceivedByEmailAddress", 0x0076);
        /// <summary>PidTagReceivedRepresentingAddressType (0x0077).</summary>
        public static readonly MapiPropertyKey<string> ReceivedRepresentingAddressType =
            String("PidTagReceivedRepresentingAddressType", 0x0077);
        /// <summary>PidTagReceivedRepresentingEmailAddress (0x0078).</summary>
        public static readonly MapiPropertyKey<string> ReceivedRepresentingEmailAddress =
            String("PidTagReceivedRepresentingEmailAddress", 0x0078);
        /// <summary>PidTagToDoItemFlags (0x0E2B).</summary>
        public static readonly MapiPropertyKey<int> ToDoItemFlags = Integer("PidTagToDoItemFlags", 0x0E2B);
        /// <summary>PidTagFlagStatus (0x1090).</summary>
        public static readonly MapiPropertyKey<int> FlagStatus = Integer("PidTagFlagStatus", 0x1090);
        /// <summary>PidTagFlagCompleteTime (0x1091).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> FlagCompleteTime = Time("PidTagFlagCompleteTime", 0x1091);
        /// <summary>PidTagFollowupIcon (0x1095).</summary>
        public static readonly MapiPropertyKey<int> FollowupIcon = Integer("PidTagFollowupIcon", 0x1095);

        /// <summary>PidTagCallbackTelephoneNumber (0x3A02).</summary>
        public static readonly MapiPropertyKey<string> CallbackTelephoneNumber = String("PidTagCallbackTelephoneNumber", 0x3A02);
        /// <summary>PidTagGeneration (0x3A05).</summary>
        public static readonly MapiPropertyKey<string> Generation = String("PidTagGeneration", 0x3A05);
        /// <summary>PidTagGivenName (0x3A06).</summary>
        public static readonly MapiPropertyKey<string> GivenName = String("PidTagGivenName", 0x3A06);
        /// <summary>PidTagBusinessTelephoneNumber (0x3A08).</summary>
        public static readonly MapiPropertyKey<string> BusinessTelephoneNumber = String("PidTagBusinessTelephoneNumber", 0x3A08);
        /// <summary>PidTagHomeTelephoneNumber (0x3A09).</summary>
        public static readonly MapiPropertyKey<string> HomeTelephoneNumber = String("PidTagHomeTelephoneNumber", 0x3A09);
        /// <summary>PidTagInitials (0x3A0A).</summary>
        public static readonly MapiPropertyKey<string> Initials = String("PidTagInitials", 0x3A0A);
        /// <summary>PidTagLanguage (0x3A0C).</summary>
        public static readonly MapiPropertyKey<string> Language = String("PidTagLanguage", 0x3A0C);
        /// <summary>PidTagLocation (0x3A0D).</summary>
        public static readonly MapiPropertyKey<string> MailUserLocation = String("PidTagLocation", 0x3A0D);
        /// <summary>PidTagSurname (0x3A11).</summary>
        public static readonly MapiPropertyKey<string> Surname = String("PidTagSurname", 0x3A11);
        /// <summary>PidTagPostalAddress (0x3A15).</summary>
        public static readonly MapiPropertyKey<string> PostalAddress = String("PidTagPostalAddress", 0x3A15);
        /// <summary>PidTagCompanyName (0x3A16).</summary>
        public static readonly MapiPropertyKey<string> CompanyName = String("PidTagCompanyName", 0x3A16);
        /// <summary>PidTagTitle (0x3A17).</summary>
        public static readonly MapiPropertyKey<string> Title = String("PidTagTitle", 0x3A17);
        /// <summary>PidTagDepartmentName (0x3A18).</summary>
        public static readonly MapiPropertyKey<string> DepartmentName = String("PidTagDepartmentName", 0x3A18);
        /// <summary>PidTagOfficeLocation (0x3A19).</summary>
        public static readonly MapiPropertyKey<string> OfficeLocation = String("PidTagOfficeLocation", 0x3A19);
        /// <summary>PidTagPrimaryTelephoneNumber (0x3A1A).</summary>
        public static readonly MapiPropertyKey<string> PrimaryTelephoneNumber = String("PidTagPrimaryTelephoneNumber", 0x3A1A);
        /// <summary>PidTagBusiness2TelephoneNumber (0x3A1B).</summary>
        public static readonly MapiPropertyKey<string> Business2TelephoneNumber = String("PidTagBusiness2TelephoneNumber", 0x3A1B);
        /// <summary>PidTagMobileTelephoneNumber (0x3A1C).</summary>
        public static readonly MapiPropertyKey<string> MobileTelephoneNumber = String("PidTagMobileTelephoneNumber", 0x3A1C);
        /// <summary>PidTagRadioTelephoneNumber (0x3A1D).</summary>
        public static readonly MapiPropertyKey<string> RadioTelephoneNumber = String("PidTagRadioTelephoneNumber", 0x3A1D);
        /// <summary>PidTagCarTelephoneNumber (0x3A1E).</summary>
        public static readonly MapiPropertyKey<string> CarTelephoneNumber = String("PidTagCarTelephoneNumber", 0x3A1E);
        /// <summary>PidTagOtherTelephoneNumber (0x3A1F).</summary>
        public static readonly MapiPropertyKey<string> OtherTelephoneNumber = String("PidTagOtherTelephoneNumber", 0x3A1F);
        /// <summary>PidTagPagerTelephoneNumber (0x3A21).</summary>
        public static readonly MapiPropertyKey<string> PagerTelephoneNumber = String("PidTagPagerTelephoneNumber", 0x3A21);
        /// <summary>PidTagPrimaryFaxNumber (0x3A23).</summary>
        public static readonly MapiPropertyKey<string> PrimaryFaxNumber = String("PidTagPrimaryFaxNumber", 0x3A23);
        /// <summary>PidTagBusinessFaxNumber (0x3A24).</summary>
        public static readonly MapiPropertyKey<string> BusinessFaxNumber = String("PidTagBusinessFaxNumber", 0x3A24);
        /// <summary>PidTagHomeFaxNumber (0x3A25).</summary>
        public static readonly MapiPropertyKey<string> HomeFaxNumber = String("PidTagHomeFaxNumber", 0x3A25);
        /// <summary>PidTagCountry (0x3A26).</summary>
        public static readonly MapiPropertyKey<string> Country = String("PidTagCountry", 0x3A26);
        /// <summary>PidTagLocality (0x3A27).</summary>
        public static readonly MapiPropertyKey<string> Locality = String("PidTagLocality", 0x3A27);
        /// <summary>PidTagStateOrProvince (0x3A28).</summary>
        public static readonly MapiPropertyKey<string> StateOrProvince = String("PidTagStateOrProvince", 0x3A28);
        /// <summary>PidTagStreetAddress (0x3A29).</summary>
        public static readonly MapiPropertyKey<string> StreetAddress = String("PidTagStreetAddress", 0x3A29);
        /// <summary>PidTagPostalCode (0x3A2A).</summary>
        public static readonly MapiPropertyKey<string> PostalCode = String("PidTagPostalCode", 0x3A2A);
        /// <summary>PidTagPostOfficeBox (0x3A2B).</summary>
        public static readonly MapiPropertyKey<string> PostOfficeBox = String("PidTagPostOfficeBox", 0x3A2B);
        /// <summary>PidTagTelexNumber (0x3A2C).</summary>
        public static readonly MapiPropertyKey<string> TelexNumber = String("PidTagTelexNumber", 0x3A2C);
        /// <summary>PidTagIsdnNumber (0x3A2D).</summary>
        public static readonly MapiPropertyKey<string> IsdnNumber = String("PidTagIsdnNumber", 0x3A2D);
        /// <summary>PidTagAssistantTelephoneNumber (0x3A2E).</summary>
        public static readonly MapiPropertyKey<string> AssistantTelephoneNumber = String("PidTagAssistantTelephoneNumber", 0x3A2E);
        /// <summary>PidTagHome2TelephoneNumber (0x3A2F).</summary>
        public static readonly MapiPropertyKey<string> Home2TelephoneNumber = String("PidTagHome2TelephoneNumber", 0x3A2F);
        /// <summary>PidTagAssistant (0x3A30).</summary>
        public static readonly MapiPropertyKey<string> Assistant = String("PidTagAssistant", 0x3A30);
        /// <summary>PidTagWeddingAnniversary (0x3A41).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> WeddingAnniversary = Time("PidTagWeddingAnniversary", 0x3A41);
        /// <summary>PidTagBirthday (0x3A42).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> Birthday = Time("PidTagBirthday", 0x3A42);
        /// <summary>PidTagMiddleName (0x3A44).</summary>
        public static readonly MapiPropertyKey<string> MiddleName = String("PidTagMiddleName", 0x3A44);
        /// <summary>PidTagDisplayNamePrefix (0x3A45).</summary>
        public static readonly MapiPropertyKey<string> DisplayNamePrefix = String("PidTagDisplayNamePrefix", 0x3A45);
        /// <summary>PidTagProfession (0x3A46).</summary>
        public static readonly MapiPropertyKey<string> Profession = String("PidTagProfession", 0x3A46);
        /// <summary>PidTagSpouseName (0x3A48).</summary>
        public static readonly MapiPropertyKey<string> SpouseName = String("PidTagSpouseName", 0x3A48);
        /// <summary>PidTagTtyTddPhoneNumber (0x3A4B).</summary>
        public static readonly MapiPropertyKey<string> TtyTddPhoneNumber = String("PidTagTtyTddPhoneNumber", 0x3A4B);
        /// <summary>PidTagManagerName (0x3A4E).</summary>
        public static readonly MapiPropertyKey<string> ManagerName = String("PidTagManagerName", 0x3A4E);
        /// <summary>PidTagNickname (0x3A4F).</summary>
        public static readonly MapiPropertyKey<string> Nickname = String("PidTagNickname", 0x3A4F);
        /// <summary>PidTagPersonalHomePage (0x3A50).</summary>
        public static readonly MapiPropertyKey<string> PersonalHomePage = String("PidTagPersonalHomePage", 0x3A50);
        /// <summary>PidTagBusinessHomePage (0x3A51).</summary>
        public static readonly MapiPropertyKey<string> BusinessHomePage = String("PidTagBusinessHomePage", 0x3A51);
        /// <summary>PidTagCompanyMainPhoneNumber (0x3A57).</summary>
        public static readonly MapiPropertyKey<string> CompanyMainPhoneNumber = String("PidTagCompanyMainPhoneNumber", 0x3A57);
        /// <summary>PidTagChildrensNames (0x3A58), accepting legacy scalar and canonical multi-value forms.</summary>
        public static readonly MapiPropertyKey<object> ChildrensNames = new MapiPropertyKey<object>("PidTagChildrensNames", 0x3A58,
            MapiPropertyType.MultipleUnicode, MapiPropertyType.MultipleString8, MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidTagHomeAddressCity (0x3A59).</summary>
        public static readonly MapiPropertyKey<string> HomeAddressCity = String("PidTagHomeAddressCity", 0x3A59);
        /// <summary>PidTagHomeAddressCountry (0x3A5A).</summary>
        public static readonly MapiPropertyKey<string> HomeAddressCountry = String("PidTagHomeAddressCountry", 0x3A5A);
        /// <summary>PidTagHomeAddressPostalCode (0x3A5B).</summary>
        public static readonly MapiPropertyKey<string> HomeAddressPostalCode = String("PidTagHomeAddressPostalCode", 0x3A5B);
        /// <summary>PidTagHomeAddressStateOrProvince (0x3A5C).</summary>
        public static readonly MapiPropertyKey<string> HomeAddressStateOrProvince = String("PidTagHomeAddressStateOrProvince", 0x3A5C);
        /// <summary>PidTagHomeAddressStreet (0x3A5D).</summary>
        public static readonly MapiPropertyKey<string> HomeAddressStreet = String("PidTagHomeAddressStreet", 0x3A5D);
        /// <summary>PidTagHomeAddressPostOfficeBox (0x3A5E).</summary>
        public static readonly MapiPropertyKey<string> HomeAddressPostOfficeBox = String("PidTagHomeAddressPostOfficeBox", 0x3A5E);
        /// <summary>PidTagOtherAddressCity (0x3A5F).</summary>
        public static readonly MapiPropertyKey<string> OtherAddressCity = String("PidTagOtherAddressCity", 0x3A5F);
        /// <summary>PidTagOtherAddressCountry (0x3A60).</summary>
        public static readonly MapiPropertyKey<string> OtherAddressCountry = String("PidTagOtherAddressCountry", 0x3A60);
        /// <summary>PidTagOtherAddressPostalCode (0x3A61).</summary>
        public static readonly MapiPropertyKey<string> OtherAddressPostalCode = String("PidTagOtherAddressPostalCode", 0x3A61);
        /// <summary>PidTagOtherAddressStateOrProvince (0x3A62).</summary>
        public static readonly MapiPropertyKey<string> OtherAddressStateOrProvince = String("PidTagOtherAddressStateOrProvince", 0x3A62);
        /// <summary>PidTagOtherAddressStreet (0x3A63).</summary>
        public static readonly MapiPropertyKey<string> OtherAddressStreet = String("PidTagOtherAddressStreet", 0x3A63);
        /// <summary>PidTagOtherAddressPostOfficeBox (0x3A64).</summary>
        public static readonly MapiPropertyKey<string> OtherAddressPostOfficeBox = String("PidTagOtherAddressPostOfficeBox", 0x3A64);
        /// <summary>PR_ORG_EMAIL_ADDR (0x403E), a transport-provider originator address fallback.</summary>
        public static readonly MapiPropertyKey<string> OriginatorEmailAddress = String("PrOrgEmailAddress", 0x403E);

        internal static readonly IReadOnlyList<MapiPropertyKey> OutlookTaggedProperties = new MapiPropertyKey[] {
            ReceivedByEntryId, ReceivedByName, ReceivedRepresentingEntryId, ReceivedRepresentingName,
            ReceivedByAddressType, ReceivedByEmailAddress, ReceivedRepresentingAddressType,
            ReceivedRepresentingEmailAddress, ToDoItemFlags, FlagStatus, FlagCompleteTime, FollowupIcon,
            CallbackTelephoneNumber, Generation, GivenName,
            BusinessTelephoneNumber, HomeTelephoneNumber, Initials, Language, MailUserLocation, Surname,
            PostalAddress, CompanyName, Title, DepartmentName, OfficeLocation, PrimaryTelephoneNumber,
            Business2TelephoneNumber, MobileTelephoneNumber, RadioTelephoneNumber, CarTelephoneNumber,
            OtherTelephoneNumber, PagerTelephoneNumber, PrimaryFaxNumber, BusinessFaxNumber, HomeFaxNumber,
            Country, Locality, StateOrProvince, StreetAddress, PostalCode, PostOfficeBox, TelexNumber, IsdnNumber,
            AssistantTelephoneNumber, Home2TelephoneNumber, Assistant, WeddingAnniversary, Birthday, MiddleName,
            DisplayNamePrefix, Profession, SpouseName, TtyTddPhoneNumber, ManagerName, Nickname, PersonalHomePage,
            BusinessHomePage, CompanyMainPhoneNumber, ChildrensNames, HomeAddressCity, HomeAddressCountry,
            HomeAddressPostalCode, HomeAddressStateOrProvince, HomeAddressStreet, HomeAddressPostOfficeBox,
            OtherAddressCity, OtherAddressCountry, OtherAddressPostalCode, OtherAddressStateOrProvince,
            OtherAddressStreet, OtherAddressPostOfficeBox, OriginatorEmailAddress
        };
    }

    public static partial class PidLid {
        /// <summary>Legacy numeric PS_PUBLIC_STRINGS Keywords identity accepted for older artifacts.</summary>
        public static readonly MapiPropertyKey<object> LegacyKeywords = new MapiPropertyKey<object>(
            "LegacyKeywords", MapiPropertySets.PublicStrings, 0x9000, MapiPropertyType.MultipleUnicode,
            MapiPropertyType.MultipleString8, MapiPropertyType.Unicode, MapiPropertyType.String8);
        /// <summary>PidLidFileUnder (0x8005).</summary>
        public static readonly MapiPropertyKey<string> FileUnder = NamedString("PidLidFileUnder", MapiPropertySets.Address, 0x8005);
        /// <summary>PidLidHasPicture (0x8015).</summary>
        public static readonly MapiPropertyKey<bool> HasPicture = NamedBoolean("PidLidHasPicture", MapiPropertySets.Address, 0x8015);
        /// <summary>PidLidHomeAddress (0x801A).</summary>
        public static readonly MapiPropertyKey<string> HomeAddress = NamedString("PidLidHomeAddress", MapiPropertySets.Address, 0x801A);
        /// <summary>PidLidWorkAddress (0x801B).</summary>
        public static readonly MapiPropertyKey<string> WorkAddress = NamedString("PidLidWorkAddress", MapiPropertySets.Address, 0x801B);
        /// <summary>PidLidOtherAddress (0x801C).</summary>
        public static readonly MapiPropertyKey<string> OtherAddress = NamedString("PidLidOtherAddress", MapiPropertySets.Address, 0x801C);
        /// <summary>PidLidHtml (0x802B).</summary>
        public static readonly MapiPropertyKey<string> ContactHtml = NamedString("PidLidHtml", MapiPropertySets.Address, 0x802B);
        /// <summary>PidLidWorkAddressStreet (0x8045).</summary>
        public static readonly MapiPropertyKey<string> WorkAddressStreet = NamedString("PidLidWorkAddressStreet", MapiPropertySets.Address, 0x8045);
        /// <summary>PidLidWorkAddressCity (0x8046).</summary>
        public static readonly MapiPropertyKey<string> WorkAddressCity = NamedString("PidLidWorkAddressCity", MapiPropertySets.Address, 0x8046);
        /// <summary>PidLidWorkAddressState (0x8047).</summary>
        public static readonly MapiPropertyKey<string> WorkAddressState = NamedString("PidLidWorkAddressState", MapiPropertySets.Address, 0x8047);
        /// <summary>PidLidWorkAddressPostalCode (0x8048).</summary>
        public static readonly MapiPropertyKey<string> WorkAddressPostalCode = NamedString("PidLidWorkAddressPostalCode", MapiPropertySets.Address, 0x8048);
        /// <summary>PidLidWorkAddressCountry (0x8049).</summary>
        public static readonly MapiPropertyKey<string> WorkAddressCountry = NamedString("PidLidWorkAddressCountry", MapiPropertySets.Address, 0x8049);
        /// <summary>PidLidWorkAddressPostOfficeBox (0x804A).</summary>
        public static readonly MapiPropertyKey<string> WorkAddressPostOfficeBox = NamedString("PidLidWorkAddressPostOfficeBox", MapiPropertySets.Address, 0x804A);
        /// <summary>PidLidInstantMessagingAddress (0x8062).</summary>
        public static readonly MapiPropertyKey<string> InstantMessagingAddress = NamedString("PidLidInstantMessagingAddress", MapiPropertySets.Address, 0x8062);
        /// <summary>PidLidEmail1DisplayName (0x8080).</summary>
        public static readonly MapiPropertyKey<string> Email1DisplayName = NamedString("PidLidEmail1DisplayName", MapiPropertySets.Address, 0x8080);
        /// <summary>PidLidEmail1AddressType (0x8082).</summary>
        public static readonly MapiPropertyKey<string> Email1AddressType = NamedString("PidLidEmail1AddressType", MapiPropertySets.Address, 0x8082);
        /// <summary>PidLidEmail1EmailAddress (0x8083).</summary>
        public static readonly MapiPropertyKey<string> Email1EmailAddress = NamedString("PidLidEmail1EmailAddress", MapiPropertySets.Address, 0x8083);
        /// <summary>PidLidEmail1OriginalDisplayName (0x8084).</summary>
        public static readonly MapiPropertyKey<string> Email1OriginalDisplayName = NamedString("PidLidEmail1OriginalDisplayName", MapiPropertySets.Address, 0x8084);
        /// <summary>PidLidEmail1OriginalEntryId (0x8085).</summary>
        public static readonly MapiPropertyKey<byte[]> Email1OriginalEntryId = NamedBinary("PidLidEmail1OriginalEntryId", MapiPropertySets.Address, 0x8085);
        /// <summary>PidLidEmail2DisplayName (0x8090).</summary>
        public static readonly MapiPropertyKey<string> Email2DisplayName = NamedString("PidLidEmail2DisplayName", MapiPropertySets.Address, 0x8090);
        /// <summary>PidLidEmail2AddressType (0x8092).</summary>
        public static readonly MapiPropertyKey<string> Email2AddressType = NamedString("PidLidEmail2AddressType", MapiPropertySets.Address, 0x8092);
        /// <summary>PidLidEmail2EmailAddress (0x8093).</summary>
        public static readonly MapiPropertyKey<string> Email2EmailAddress = NamedString("PidLidEmail2EmailAddress", MapiPropertySets.Address, 0x8093);
        /// <summary>PidLidEmail2OriginalDisplayName (0x8094).</summary>
        public static readonly MapiPropertyKey<string> Email2OriginalDisplayName = NamedString("PidLidEmail2OriginalDisplayName", MapiPropertySets.Address, 0x8094);
        /// <summary>PidLidEmail2OriginalEntryId (0x8095).</summary>
        public static readonly MapiPropertyKey<byte[]> Email2OriginalEntryId = NamedBinary("PidLidEmail2OriginalEntryId", MapiPropertySets.Address, 0x8095);
        /// <summary>PidLidEmail3DisplayName (0x80A0).</summary>
        public static readonly MapiPropertyKey<string> Email3DisplayName = NamedString("PidLidEmail3DisplayName", MapiPropertySets.Address, 0x80A0);
        /// <summary>PidLidEmail3AddressType (0x80A2).</summary>
        public static readonly MapiPropertyKey<string> Email3AddressType = NamedString("PidLidEmail3AddressType", MapiPropertySets.Address, 0x80A2);
        /// <summary>PidLidEmail3EmailAddress (0x80A3).</summary>
        public static readonly MapiPropertyKey<string> Email3EmailAddress = NamedString("PidLidEmail3EmailAddress", MapiPropertySets.Address, 0x80A3);
        /// <summary>PidLidEmail3OriginalDisplayName (0x80A4).</summary>
        public static readonly MapiPropertyKey<string> Email3OriginalDisplayName = NamedString("PidLidEmail3OriginalDisplayName", MapiPropertySets.Address, 0x80A4);
        /// <summary>PidLidEmail3OriginalEntryId (0x80A5).</summary>
        public static readonly MapiPropertyKey<byte[]> Email3OriginalEntryId = NamedBinary("PidLidEmail3OriginalEntryId", MapiPropertySets.Address, 0x80A5);
        /// <summary>PidLidWorkAddressCountryCode (0x80DB).</summary>
        public static readonly MapiPropertyKey<string> WorkAddressCountryCode = NamedString("PidLidWorkAddressCountryCode", MapiPropertySets.Address, 0x80DB);
        /// <summary>PidLidBirthdayLocal (0x80DE).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> BirthdayLocal = NamedTime("PidLidBirthdayLocal", MapiPropertySets.Address, 0x80DE);
        /// <summary>PidLidWeddingAnniversaryLocal (0x80DF).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> WeddingAnniversaryLocal = NamedTime("PidLidWeddingAnniversaryLocal", MapiPropertySets.Address, 0x80DF);

        /// <summary>PidLidDistributionListChecksum (0x804C).</summary>
        public static readonly MapiPropertyKey<int> DistributionListChecksum =
            NamedInteger("PidLidDistributionListChecksum", MapiPropertySets.Address, 0x804C);
        /// <summary>PidLidDistributionListName (0x8053).</summary>
        public static readonly MapiPropertyKey<string> DistributionListName =
            NamedString("PidLidDistributionListName", MapiPropertySets.Address, 0x8053);
        /// <summary>PidLidDistributionListOneOffMembers (0x8054).</summary>
        public static readonly MapiPropertyKey<object[]> DistributionListOneOffMembers =
            NamedMultipleBinary("PidLidDistributionListOneOffMembers", MapiPropertySets.Address, 0x8054);
        /// <summary>PidLidDistributionListMembers (0x8055).</summary>
        public static readonly MapiPropertyKey<object[]> DistributionListMembers =
            NamedMultipleBinary("PidLidDistributionListMembers", MapiPropertySets.Address, 0x8055);

        /// <summary>PidLidTaskStatus (0x8101).</summary>
        public static readonly MapiPropertyKey<int> TaskStatus = NamedInteger("PidLidTaskStatus", MapiPropertySets.Task, 0x8101);
        /// <summary>PidLidPercentComplete (0x8102).</summary>
        public static readonly MapiPropertyKey<double> PercentComplete = new MapiPropertyKey<double>("PidLidPercentComplete", MapiPropertySets.Task, 0x8102, MapiPropertyType.Floating64, MapiPropertyType.Floating32);
        /// <summary>PidLidTeamTask (0x8103).</summary>
        public static readonly MapiPropertyKey<bool> TeamTask = NamedBoolean("PidLidTeamTask", MapiPropertySets.Task, 0x8103);
        /// <summary>PidLidTaskStartDate (0x8104).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> TaskStartDate = NamedTime("PidLidTaskStartDate", MapiPropertySets.Task, 0x8104);
        /// <summary>PidLidTaskDueDate (0x8105).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> TaskDueDate = NamedTime("PidLidTaskDueDate", MapiPropertySets.Task, 0x8105);
        /// <summary>PidLidTaskAccepted (0x8108).</summary>
        public static readonly MapiPropertyKey<bool> TaskAccepted = NamedBoolean("PidLidTaskAccepted", MapiPropertySets.Task, 0x8108);
        /// <summary>PidLidTaskDateCompleted (0x810F).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> TaskDateCompleted =
            NamedTime("PidLidTaskDateCompleted", MapiPropertySets.Task, 0x810F);
        /// <summary>PidLidTaskActualEffort (0x8110).</summary>
        public static readonly MapiPropertyKey<int> TaskActualEffort = NamedInteger("PidLidTaskActualEffort", MapiPropertySets.Task, 0x8110);
        /// <summary>PidLidTaskEstimatedEffort (0x8111).</summary>
        public static readonly MapiPropertyKey<int> TaskEstimatedEffort = NamedInteger("PidLidTaskEstimatedEffort", MapiPropertySets.Task, 0x8111);
        /// <summary>PidLidTaskVersion (0x8112).</summary>
        public static readonly MapiPropertyKey<int> TaskVersion = NamedInteger("PidLidTaskVersion", MapiPropertySets.Task, 0x8112);
        /// <summary>PidLidTaskState (0x8113).</summary>
        public static readonly MapiPropertyKey<int> TaskState = NamedInteger("PidLidTaskState", MapiPropertySets.Task, 0x8113);
        /// <summary>PidLidTaskLastUpdate (0x8115).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> TaskLastUpdate = NamedTime("PidLidTaskLastUpdate", MapiPropertySets.Task, 0x8115);
        /// <summary>PidLidTaskRecurrence (0x8116).</summary>
        public static readonly MapiPropertyKey<byte[]> TaskRecurrence =
            NamedBinary("PidLidTaskRecurrence", MapiPropertySets.Task, 0x8116);
        /// <summary>PidLidTaskStatusOnComplete (0x8119).</summary>
        public static readonly MapiPropertyKey<bool> TaskStatusOnComplete = NamedBoolean("PidLidTaskStatusOnComplete", MapiPropertySets.Task, 0x8119);
        /// <summary>PidLidTaskHistory (0x811A).</summary>
        public static readonly MapiPropertyKey<int> TaskHistory = NamedInteger("PidLidTaskHistory", MapiPropertySets.Task, 0x811A);
        /// <summary>PidLidTaskUpdates (0x811B).</summary>
        public static readonly MapiPropertyKey<bool> TaskUpdates = NamedBoolean("PidLidTaskUpdates", MapiPropertySets.Task, 0x811B);
        /// <summary>PidLidTaskComplete (0x811C).</summary>
        public static readonly MapiPropertyKey<bool> TaskComplete = NamedBoolean("PidLidTaskComplete", MapiPropertySets.Task, 0x811C);
        /// <summary>PidLidTaskOwner (0x811F).</summary>
        public static readonly MapiPropertyKey<string> TaskOwner = NamedString("PidLidTaskOwner", MapiPropertySets.Task, 0x811F);
        /// <summary>PidLidTaskLastUser (0x8122).</summary>
        public static readonly MapiPropertyKey<string> TaskLastUser = NamedString("PidLidTaskLastUser", MapiPropertySets.Task, 0x8122);
        /// <summary>PidLidTaskAssigner (0x8121).</summary>
        public static readonly MapiPropertyKey<string> TaskAssigner = NamedString("PidLidTaskAssigner", MapiPropertySets.Task, 0x8121);
        /// <summary>PidLidTaskOrdinal (0x8123).</summary>
        public static readonly MapiPropertyKey<int> TaskOrdinal = NamedInteger("PidLidTaskOrdinal", MapiPropertySets.Task, 0x8123);
        /// <summary>PidLidTaskLastDelegate (0x8125).</summary>
        public static readonly MapiPropertyKey<string> TaskLastDelegate = NamedString("PidLidTaskLastDelegate", MapiPropertySets.Task, 0x8125);
        /// <summary>PidLidTaskFRecurring (0x8126).</summary>
        public static readonly MapiPropertyKey<bool> TaskFRecurring =
            NamedBoolean("PidLidTaskFRecurring", MapiPropertySets.Task, 0x8126);
        /// <summary>PidLidTaskOwnership (0x8129).</summary>
        public static readonly MapiPropertyKey<int> TaskOwnership = NamedInteger("PidLidTaskOwnership", MapiPropertySets.Task, 0x8129);
        /// <summary>PidLidTaskAcceptanceState (0x812A).</summary>
        public static readonly MapiPropertyKey<int> TaskAcceptanceState = NamedInteger("PidLidTaskAcceptanceState", MapiPropertySets.Task, 0x812A);
        /// <summary>PidLidTaskMode (0x8518).</summary>
        public static readonly MapiPropertyKey<int> TaskMode = NamedInteger("PidLidTaskMode", MapiPropertySets.Common, 0x8518);
        /// <summary>PidLidTaskGlobalId (0x8519).</summary>
        public static readonly MapiPropertyKey<byte[]> TaskGlobalId = NamedBinary("PidLidTaskGlobalId", MapiPropertySets.Common, 0x8519);
        /// <summary>PidLidMileage (0x8534).</summary>
        public static readonly MapiPropertyKey<string> Mileage = NamedString("PidLidMileage", MapiPropertySets.Common, 0x8534);
        /// <summary>PidLidBilling (0x8535).</summary>
        public static readonly MapiPropertyKey<string> Billing = NamedString("PidLidBilling", MapiPropertySets.Common, 0x8535);
        /// <summary>PidLidCompanies (0x8539).</summary>
        public static readonly MapiPropertyKey<object[]> Companies = NamedMultipleString("PidLidCompanies", MapiPropertySets.Common, 0x8539);
        /// <summary>PidLidContacts (0x853A).</summary>
        public static readonly MapiPropertyKey<object[]> Contacts = NamedMultipleString("PidLidContacts", MapiPropertySets.Common, 0x853A);
        /// <summary>PidLidToDoOrdinalDate (0x85A0).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> ToDoOrdinalDate = NamedTime("PidLidToDoOrdinalDate", MapiPropertySets.Common, 0x85A0);
        /// <summary>PidLidToDoSubOrdinal (0x85A1).</summary>
        public static readonly MapiPropertyKey<string> ToDoSubOrdinal = NamedString("PidLidToDoSubOrdinal", MapiPropertySets.Common, 0x85A1);

        /// <summary>PidLidLogType (0x8700).</summary>
        public static readonly MapiPropertyKey<string> LogType = NamedString("PidLidLogType", MapiPropertySets.Log, 0x8700);
        /// <summary>PidLidLogStart (0x8706).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> LogStart = NamedTime("PidLidLogStart", MapiPropertySets.Log, 0x8706);
        /// <summary>PidLidLogDuration (0x8707).</summary>
        public static readonly MapiPropertyKey<int> LogDuration = NamedInteger("PidLidLogDuration", MapiPropertySets.Log, 0x8707);
        /// <summary>PidLidLogEnd (0x8708).</summary>
        public static readonly MapiPropertyKey<DateTimeOffset> LogEnd = NamedTime("PidLidLogEnd", MapiPropertySets.Log, 0x8708);
        /// <summary>PidLidLogFlags (0x870C).</summary>
        public static readonly MapiPropertyKey<int> LogFlags = NamedInteger("PidLidLogFlags", MapiPropertySets.Log, 0x870C);
        /// <summary>PidLidLogDocumentPrinted (0x870E).</summary>
        public static readonly MapiPropertyKey<bool> LogDocumentPrinted = NamedBoolean("PidLidLogDocumentPrinted", MapiPropertySets.Log, 0x870E);
        /// <summary>PidLidLogDocumentSaved (0x870F).</summary>
        public static readonly MapiPropertyKey<bool> LogDocumentSaved = NamedBoolean("PidLidLogDocumentSaved", MapiPropertySets.Log, 0x870F);
        /// <summary>PidLidLogDocumentRouted (0x8710).</summary>
        public static readonly MapiPropertyKey<bool> LogDocumentRouted = NamedBoolean("PidLidLogDocumentRouted", MapiPropertySets.Log, 0x8710);
        /// <summary>PidLidLogDocumentPosted (0x8711).</summary>
        public static readonly MapiPropertyKey<bool> LogDocumentPosted = NamedBoolean("PidLidLogDocumentPosted", MapiPropertySets.Log, 0x8711);
        /// <summary>PidLidLogTypeDesc (0x8712).</summary>
        public static readonly MapiPropertyKey<string> LogTypeDesc = NamedString("PidLidLogTypeDesc", MapiPropertySets.Log, 0x8712);

        /// <summary>PidLidNoteColor (0x8B00).</summary>
        public static readonly MapiPropertyKey<int> NoteColor = NamedInteger("PidLidNoteColor", MapiPropertySets.Note, 0x8B00);
        /// <summary>PidLidNoteWidth (0x8B02).</summary>
        public static readonly MapiPropertyKey<int> NoteWidth = NamedInteger("PidLidNoteWidth", MapiPropertySets.Note, 0x8B02);
        /// <summary>PidLidNoteHeight (0x8B03).</summary>
        public static readonly MapiPropertyKey<int> NoteHeight = NamedInteger("PidLidNoteHeight", MapiPropertySets.Note, 0x8B03);
        /// <summary>PidLidNoteX (0x8B04).</summary>
        public static readonly MapiPropertyKey<int> NoteX = NamedInteger("PidLidNoteX", MapiPropertySets.Note, 0x8B04);
        /// <summary>PidLidNoteY (0x8B05).</summary>
        public static readonly MapiPropertyKey<int> NoteY = NamedInteger("PidLidNoteY", MapiPropertySets.Note, 0x8B05);

        internal static readonly IReadOnlyList<MapiPropertyKey> OutlookNamedProperties = new MapiPropertyKey[] {
            LegacyKeywords, FileUnder, HasPicture, HomeAddress, WorkAddress, OtherAddress, ContactHtml, WorkAddressStreet,
            WorkAddressCity, WorkAddressState, WorkAddressPostalCode, WorkAddressCountry, WorkAddressPostOfficeBox,
            InstantMessagingAddress, Email1DisplayName, Email1AddressType, Email1EmailAddress,
            Email1OriginalDisplayName, Email1OriginalEntryId, Email2DisplayName, Email2AddressType,
            Email2EmailAddress, Email2OriginalDisplayName, Email2OriginalEntryId, Email3DisplayName,
            Email3AddressType, Email3EmailAddress, Email3OriginalDisplayName, Email3OriginalEntryId,
            WorkAddressCountryCode, BirthdayLocal, WeddingAnniversaryLocal, TaskStatus, PercentComplete, TeamTask,
            DistributionListChecksum, DistributionListName, DistributionListOneOffMembers, DistributionListMembers,
            TaskStartDate, TaskDueDate, TaskAccepted, TaskDateCompleted, TaskActualEffort, TaskEstimatedEffort,
            TaskVersion, TaskState, TaskLastUpdate, TaskRecurrence, TaskStatusOnComplete, TaskHistory, TaskUpdates,
            TaskComplete, TaskOwner, TaskLastUser, TaskAssigner, TaskOrdinal, TaskLastDelegate, TaskOwnership,
            TaskFRecurring, TaskAcceptanceState, TaskMode, TaskGlobalId, Mileage, Billing, Companies, Contacts, ToDoOrdinalDate,
            ToDoSubOrdinal,
            LogType, LogStart, LogDuration, LogEnd, LogFlags, LogDocumentPrinted, LogDocumentSaved,
            LogDocumentRouted, LogDocumentPosted, LogTypeDesc, NoteColor, NoteWidth, NoteHeight, NoteX, NoteY
        };

        private static MapiPropertyKey<string> NamedString(string name, Guid propertySet, uint localId) =>
            new MapiPropertyKey<string>(name, propertySet, localId, MapiPropertyType.Unicode, MapiPropertyType.String8);

        private static MapiPropertyKey<int> NamedInteger(string name, Guid propertySet, uint localId) =>
            new MapiPropertyKey<int>(name, propertySet, localId, MapiPropertyType.Integer32,
                MapiPropertyType.Integer16, MapiPropertyType.Integer64);

        private static MapiPropertyKey<bool> NamedBoolean(string name, Guid propertySet, uint localId) =>
            new MapiPropertyKey<bool>(name, propertySet, localId, MapiPropertyType.Boolean);

        private static MapiPropertyKey<DateTimeOffset> NamedTime(string name, Guid propertySet, uint localId) =>
            new MapiPropertyKey<DateTimeOffset>(name, propertySet, localId, MapiPropertyType.Time);

        private static MapiPropertyKey<byte[]> NamedBinary(string name, Guid propertySet, uint localId) =>
            new MapiPropertyKey<byte[]>(name, propertySet, localId, MapiPropertyType.Binary);

        private static MapiPropertyKey<object[]> NamedMultipleString(string name, Guid propertySet, uint localId) =>
            new MapiPropertyKey<object[]>(name, propertySet, localId,
                MapiPropertyType.MultipleUnicode, MapiPropertyType.MultipleString8);

        private static MapiPropertyKey<object[]> NamedMultipleBinary(string name, Guid propertySet, uint localId) =>
            new MapiPropertyKey<object[]>(name, propertySet, localId, MapiPropertyType.MultipleBinary);
    }
}
