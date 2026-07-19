using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook.Tests;

// Test-fixture aliases derive every ID from the public shared vocabulary.
internal static class OabPropertyTags {
    internal static readonly ushort DisplayName = Id(MapiKnownProperties.PidTag.DisplayName);
    internal static readonly ushort EmailAddress = Id(MapiKnownProperties.PidTag.EmailAddress);
    internal static readonly ushort ObjectType = Id(MapiKnownProperties.PidTag.ObjectType);
    internal static readonly ushort SmtpAddress = Id(MapiKnownProperties.PidTag.SmtpAddress);
    internal static readonly ushort Account = Id(MapiKnownProperties.PidTag.Account);
    internal static readonly ushort GivenName = Id(MapiKnownProperties.PidTag.GivenName);
    internal static readonly ushort BusinessTelephone = Id(MapiKnownProperties.PidTag.BusinessTelephoneNumber);
    internal static readonly ushort Surname = Id(MapiKnownProperties.PidTag.Surname);
    internal static readonly ushort CompanyName = Id(MapiKnownProperties.PidTag.CompanyName);
    internal static readonly ushort Department = Id(MapiKnownProperties.PidTag.DepartmentName);
    internal static readonly ushort SendRichInfo = Id(MapiKnownProperties.PidTag.SendRichInfo);
    internal static readonly ushort ProxyAddresses = Id(MapiKnownProperties.PidTag.AddressBookProxyAddresses);
    internal static readonly ushort Members = Id(MapiKnownProperties.PidTag.AddressBookMembers);
    internal static readonly ushort MemberOf = Id(MapiKnownProperties.PidTag.AddressBookMemberOf);
    internal static readonly ushort TruncatedProperties =
        Id(MapiKnownProperties.PidTag.OfflineAddressBookTruncatedProperties);
    internal static readonly ushort AddressBookName = Id(MapiKnownProperties.PidTag.OfflineAddressBookName);
    internal static readonly ushort AddressBookSequence = Id(MapiKnownProperties.PidTag.OfflineAddressBookSequence);
    internal static readonly ushort AddressBookContainerGuid =
        Id(MapiKnownProperties.PidTag.OfflineAddressBookContainerGuid);
    internal static readonly ushort AddressBookDistinguishedName =
        Id(MapiKnownProperties.PidTag.OfflineAddressBookDistinguishedName);

    private static ushort Id(MapiPropertyKey key) => key.GetStandardPropertyId();
}
