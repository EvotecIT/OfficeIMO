namespace OfficeIMO.Email.AddressBook;

/// <summary>How normalized search terms are combined.</summary>
public enum OfflineAddressBookSearchMatchMode {
    /// <summary>At least one term must occur in a selected field.</summary>
    AnyTerm,
    /// <summary>Every term must occur, potentially across different selected fields.</summary>
    AllTerms
}
