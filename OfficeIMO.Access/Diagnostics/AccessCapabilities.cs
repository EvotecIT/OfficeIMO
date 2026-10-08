namespace OfficeIMO.Access;

/// <summary>One operation's qualified boundary. A detected extension or profile does not imply support.</summary>
public sealed class AccessOperationCapability {
    internal AccessOperationCapability(string operation, bool supported, string boundary) { Operation = operation; IsSupported = supported; Boundary = boundary; }
    /// <summary>Stable operation identifier.</summary>
    public string Operation { get; }
    /// <summary>Whether the described operation has implemented qualification.</summary>
    public bool IsSupported { get; }
    /// <summary>Exact evidence and limitation boundary.</summary>
    public string Boundary { get; }
}

/// <summary>Canonical operation catalog used by consumers and the generated support matrix.</summary>
public static class AccessCapabilities {
    /// <summary>Current capability contracts. Native header recognition remains distinct from catalog decoding.</summary>
    public static IReadOnlyList<AccessOperationCapability> Operations { get; } = Array.AsReadOnly(new[] {
        new AccessOperationCapability("model.create", true, "New Jet4/ACE12-targeted in-memory model; no native output."),
        new AccessOperationCapability("model.edit", true, "Typed tables, columns, primary-key definitions, relationships, inert query text and rows; rollback and read leases."),
        new AccessOperationCapability("model.rows.read", true, "Forward-only DbDataReader over modeled rows; omitted values remain distinguishable from explicit nulls."),
        new AccessOperationCapability("native.header.inspect", true, "Inert bounded header/profile, page-alignment, byte/page limits and snapshot SHA-256. Independent Jet4/ACE12/ACE14/ACE16/ACE17 fixtures; Jet3 recognized without producer qualification. Protection and feature values are not decoded."),
        new AccessOperationCapability("native.catalog.read", false, "No native table, query, relationship or application catalog decoding. Loaded collections are NotDecoded."),
        new AccessOperationCapability("native.rows.read", false, "Native Jet/ACE row codecs are not qualified."),
        new AccessOperationCapability("native.create", false, "Template-free catalog/allocation/index/relationship writing is not qualified."),
        new AccessOperationCapability("native.edit", false, "Native editing and opaque preservation are not qualified."),
        new AccessOperationCapability("native.convert", false, "MDB/ACCDB feature/loss mapping is not qualified."),
        new AccessOperationCapability("application.objects.read", false, "Forms, reports and action macros are NotDecoded for native input."),
        new AccessOperationCapability("application.objects.write", false, "No native form/report/macro/module writer is qualified."),
        new AccessOperationCapability("vba.inspect", false, "Access-specific VBA storage and signature carriers are not decoded."),
        new AccessOperationCapability("protection.inspect", false, "Password, encryption and signature state is NotAssessed."),
        new AccessOperationCapability("report.render", false, "Native report definitions and expression evaluation are not qualified.")
    });
}
