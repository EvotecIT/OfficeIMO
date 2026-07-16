namespace OfficeIMO.Email.AddressBook;

internal sealed class OabAddressListSource {
    internal OabAddressListSource(OabSource source, OfflineAddressBookListInfo info) {
        Source = source;
        Info = info;
    }

    internal OabSource Source { get; }
    internal OfflineAddressBookListInfo Info { get; }
}
