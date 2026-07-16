namespace OfficeIMO.Email.AddressBook;

internal sealed class OabAddressListSource {
    internal OabAddressListSource(OabSource source, OfflineAddressBookListInfo info) {
        Source = source;
        Info = info;
        SnapshotId = Guid.NewGuid();
    }

    internal OabSource Source { get; }
    internal OfflineAddressBookListInfo Info { get; }
    internal Guid SnapshotId { get; }
}
