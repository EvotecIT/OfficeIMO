namespace OfficeIMO;

/// <summary>Adapter-owned settings that remain independent of the caller after capture.</summary>
public interface IOfficeConversionSettings {
    /// <summary>Returns an independent settings snapshot, retaining only explicitly caller-owned runtime services.</summary>
    IOfficeConversionSettings Snapshot();
}

/// <summary>An opt-in implementation of a route in the canonical conversion capability catalog.</summary>
public interface IOfficeConversionRoute {
    /// <summary>Stable identifier from <see cref="OfficeConversionCapabilityCatalog"/>.</summary>
    string Id { get; }
    /// <summary>Converts captured input to a private bounded output. Leaves both caller-owned streams open.</summary>
    Task<OfficeConversionReport> ConvertAsync(Stream source, Stream destination, OfficeConversionOptions options,
        CancellationToken cancellationToken = default);
}
