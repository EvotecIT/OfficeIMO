namespace OfficeIMO.Email.Store;

/// <summary>How normalized search terms are combined.</summary>
public enum EmailStoreContentMatchMode {
    /// <summary>At least one term must occur in a selected field.</summary>
    AnyTerm,
    /// <summary>Every term must occur, potentially across different selected fields.</summary>
    AllTerms
}
