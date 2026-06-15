namespace OfficeIMO.Rtf;

/// <summary>
/// RTF object embedding/linking kind.
/// </summary>
public enum RtfObjectKind {
    /// <summary>Object kind was not declared or is not recognized.</summary>
    Unknown,

    /// <summary>Embedded object represented by <c>\objemb</c>.</summary>
    Embedded,

    /// <summary>Linked object represented by <c>\objlink</c>.</summary>
    Linked,

    /// <summary>Automatically linked object represented by <c>\objautlink</c>.</summary>
    AutoLinked,

    /// <summary>Subscription object represented by <c>\objsub</c>.</summary>
    Subscription,

    /// <summary>Publisher object represented by <c>\objpub</c>.</summary>
    Publisher,

    /// <summary>Iconic embedded object represented by <c>\objicemb</c>.</summary>
    IconEmbedded
}
