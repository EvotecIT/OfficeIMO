namespace OfficeIMO.Data;

/// <summary>Controls how object row columns are discovered.</summary>
public enum TabularColumnDiscoveryMode {
    /// <summary>Use columns from the first projected row.</summary>
    FirstRow,

    /// <summary>Use the union of columns found across all projected rows.</summary>
    AllRows
}
