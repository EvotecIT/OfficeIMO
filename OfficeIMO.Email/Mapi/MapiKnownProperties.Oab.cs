namespace OfficeIMO.Email;

public static partial class MapiKnownProperties {
    public static partial class PidTag {
        /// <summary>PidTagComment (0x3004).</summary>
        public static readonly MapiPropertyKey<string> Comment = String("PidTagComment", 0x3004);
        /// <summary>PidTagAccount (0x3A00).</summary>
        public static readonly MapiPropertyKey<string> Account = String("PidTagAccount", 0x3A00);
        /// <summary>PidTagBusiness2TelephoneNumber in its OAB multi-value form (0x3A1B).</summary>
        public static readonly MapiPropertyKey<string[]> Business2TelephoneNumbers = new MapiPropertyKey<string[]>(
            "PidTagBusiness2TelephoneNumber", 0x3A1B, MapiPropertyType.MultipleUnicode,
            MapiPropertyType.MultipleString8);
        /// <summary>PidTagHome2TelephoneNumber in its OAB multi-value form (0x3A2F).</summary>
        public static readonly MapiPropertyKey<string[]> Home2TelephoneNumbers = new MapiPropertyKey<string[]>(
            "PidTagHome2TelephoneNumber", 0x3A2F, MapiPropertyType.MultipleUnicode,
            MapiPropertyType.MultipleString8);
        /// <summary>OAB address-list sequence property (0x6801).</summary>
        public static readonly MapiPropertyKey<int> OfflineAddressBookSequence = Integer("PidTagOfflineAddressBookSequence", 0x6801);
        /// <summary>OAB address-list container GUID text (0x6802).</summary>
        public static readonly MapiPropertyKey<string> OfflineAddressBookContainerGuid = String("PidTagOfflineAddressBookContainerGuid", 0x6802);
        /// <summary>OAB address-list distinguished name (0x6804).</summary>
        public static readonly MapiPropertyKey<string> OfflineAddressBookDistinguishedName = String("PidTagOfflineAddressBookDistinguishedName", 0x6804);
        /// <summary>Address Book member-of distinguished names (0x8008).</summary>
        public static readonly MapiPropertyKey<string[]> AddressBookMemberOf = MultipleString("PidTagAddressBookMemberOf", 0x8008);
        /// <summary>Address Book member distinguished names (0x8009).</summary>
        public static readonly MapiPropertyKey<string[]> AddressBookMembers = MultipleString("PidTagAddressBookMembers", 0x8009);
        /// <summary>Address Book proxy addresses (0x800F).</summary>
        public static readonly MapiPropertyKey<string[]> AddressBookProxyAddresses = MultipleString("PidTagAddressBookProxyAddresses", 0x800F);
        /// <summary>Address Book target address (0x8011).</summary>
        public static readonly MapiPropertyKey<string> AddressBookTargetAddress = String("PidTagAddressBookTargetAddress", 0x8011);
        /// <summary>Address Book hierarchical root department (0x8C98).</summary>
        public static readonly MapiPropertyKey<string> AddressBookHierarchicalRootDepartment =
            String("PidTagAddressBookHierarchicalRootDepartment", 0x8C98);
        /// <summary>Address Book hierarchical-group marker (0x8CDD).</summary>
        public static readonly MapiPropertyKey<bool> AddressBookHierarchicalGroup =
            Boolean("PidTagAddressBookHierarchicalGroup", 0x8CDD);
        /// <summary>Address Book distribution-list member count (0x8CE2).</summary>
        public static readonly MapiPropertyKey<int> AddressBookDistributionListMemberCount =
            Integer("PidTagAddressBookDistributionListMemberCount", 0x8CE2);
        /// <summary>Address Book external distribution-list member count (0x8CE3).</summary>
        public static readonly MapiPropertyKey<int> AddressBookDistributionListExternalMemberCount =
            Integer("PidTagAddressBookDistributionListExternalMemberCount", 0x8CE3);

        internal static readonly IReadOnlyList<MapiPropertyKey> OabTaggedProperties = new MapiPropertyKey[] {
            Comment, Account, Business2TelephoneNumbers, Home2TelephoneNumbers, OfflineAddressBookSequence,
            OfflineAddressBookContainerGuid, OfflineAddressBookDistinguishedName, AddressBookMemberOf,
            AddressBookMembers, AddressBookProxyAddresses, AddressBookTargetAddress,
            AddressBookHierarchicalRootDepartment, AddressBookHierarchicalGroup,
            AddressBookDistributionListMemberCount, AddressBookDistributionListExternalMemberCount
        };

        private static MapiPropertyKey<string[]> MultipleString(string name, ushort id) =>
            new MapiPropertyKey<string[]>(name, id, MapiPropertyType.MultipleUnicode,
                MapiPropertyType.MultipleString8);
    }
}
