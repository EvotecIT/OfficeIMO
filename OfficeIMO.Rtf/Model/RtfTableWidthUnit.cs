namespace OfficeIMO.Rtf;

/// <summary>
/// Width units used by RTF table preferred-width controls such as <c>\trftsWidth</c>.
/// </summary>
public enum RtfTableWidthUnit {
    /// <summary>Automatic width.</summary>
    Auto,

    /// <summary>Width in twips.</summary>
    Twips,

    /// <summary>Width in fiftieths of a percent, where 5000 means 100%.</summary>
    Percent
}
