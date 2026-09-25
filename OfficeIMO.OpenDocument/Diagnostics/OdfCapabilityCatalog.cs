namespace OfficeIMO.OpenDocument;

/// <summary>Implementation level for a stable OpenDocument capability line.</summary>
public enum OdfCapabilityLevel {
    /// <summary>The typed surface covers the documented capability.</summary>
    Editable,
    /// <summary>A useful, documented subset is editable while advanced variants remain preserved or unsupported.</summary>
    Limited,
    /// <summary>The feature can be detected and inspected but not edited.</summary>
    Inspected,
    /// <summary>Unchanged source content is retained without typed editing.</summary>
    Preserved,
    /// <summary>The feature is detected and rejected before unsafe or partial processing.</summary>
    DetectedUnsupported
}

/// <summary>One stable public capability declaration.</summary>
public sealed class OdfCapability {
    internal OdfCapability(string id, string name, OdfCapabilityLevel level, string description) {
        Id = id; Name = name; Level = level; Description = description;
    }
    /// <summary>Stable machine-readable identifier.</summary>
    public string Id { get; }
    /// <summary>User-facing capability name.</summary>
    public string Name { get; }
    /// <summary>Current implementation level.</summary>
    public OdfCapabilityLevel Level { get; }
    /// <summary>Concrete supported subset and boundary.</summary>
    public string Description { get; }
}

/// <summary>Stable capability catalog for advanced features that evolve independently.</summary>
public static class OdfCapabilityCatalog {
    private static readonly IReadOnlyList<OdfCapability> Capabilities = new[] {
        new OdfCapability("formula-evaluation", "OpenFormula evaluation", OdfCapabilityLevel.Limited,
            "Bounded arithmetic, comparison, concatenation, local and cross-sheet references, ranges, and common aggregate/math functions; no external data, volatile functions, or script execution."),
        new OdfCapability("conditional-style-maps", "Conditional style maps", OdfCapabilityLevel.Limited,
            "Native style maps on repository-backed common and automatic style families can be authored, edited, and preserved in order. Applied styles must be common named styles in the same family. OfficeIMO renderers do not evaluate map conditions; maps in other style families, data styles, and embedded objects are preservation-only."),
        new OdfCapability("tracked-change-editing", "Tracked-change editing", OdfCapabilityLevel.Limited,
            "Paragraph insertion and deletion changes can be authored, inspected, accepted, and rejected; arbitrary inline merge/conflict editing remains outside the typed surface."),
        new OdfCapability("advanced-charts", "Advanced charts", OdfCapabilityLevel.Limited,
            "ODS embedded charts expose structure and source ranges; bounded column, bar, and line charts can be authored from existing cell ranges. Imported chart XML is preserved, while editing imported charts and advanced chart features remain outside the typed surface."),
        new OdfCapability("data-pilot-tables", "Data pilot tables", OdfCapabilityLevel.Limited,
            "Local-range ODS data pilot tables and their source fields can be inspected and authored; advanced imported field settings remain preserved as source XML."),
        new OdfCapability("presentation-animations", "Presentation animations", OdfCapabilityLevel.Limited,
            "Basic shape attribute animations and fade-in effects can be authored and inspected; advanced timing trees remain preservation-oriented."),
        new OdfCapability("flat-xml", "Flat OpenDocument XML", OdfCapabilityLevel.Limited,
            "FODT, FODS, and FODP-style single XML documents can be opened and written, including embedded raster image binary data; package-only and exotic embedded-object features may not project losslessly."),
        new OdfCapability("encryption", "OpenDocument encryption", OdfCapabilityLevel.DetectedUnsupported,
            "Manifest encryption is detected before editing and rejected; native key derivation, decryption, and encryption are not implemented."),
        new OdfCapability("digital-signatures", "OpenDocument digital signatures", OdfCapabilityLevel.Limited,
            "Unchanged signature entries are retained and changed signed documents fail unless invalidated signatures are explicitly removed. The bounded OfficeIMO XML package-manifest profile can be created and cryptographically validated through an explicit security provider; arbitrary producer profiles remain inspection or preservation oriented.")
    };

    /// <summary>Advanced capabilities in stable catalog order.</summary>
    public static IReadOnlyList<OdfCapability> Advanced => Capabilities;

    /// <summary>Finds one capability by stable identifier.</summary>
    public static OdfCapability? Find(string id) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("Capability identifier cannot be empty.", nameof(id));
        return Capabilities.FirstOrDefault(capability => string.Equals(capability.Id, id, StringComparison.Ordinal));
    }
}
