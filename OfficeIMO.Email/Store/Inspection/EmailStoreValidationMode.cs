namespace OfficeIMO.Email.Store;

/// <summary>Depth used by bounded email-store validation.</summary>
public enum EmailStoreValidationMode {
    /// <summary>Validate only the header, indexes, folder catalog, and opening diagnostics.</summary>
    Shallow = 0,

    /// <summary>Enumerate items and selectively decode lightweight summary properties.</summary>
    Summaries = 1,

    /// <summary>Fully project each selected item, including recipients and attachment metadata or payloads.</summary>
    FullItems = 2
}
