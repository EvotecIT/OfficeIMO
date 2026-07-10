namespace OfficeIMO.Rtf;

/// <summary>Outlook/Exchange HTML encapsulated in an RTF transport.</summary>
public sealed class RtfHtmlEncapsulation {
    /// <summary>Creates an Outlook/Exchange HTML encapsulation payload.</summary>
    public RtfHtmlEncapsulation(int version, string html) {
        Version = version;
        Html = html ?? string.Empty;
    }

    /// <summary>Value declared by the RTF <c>\fromhtml</c> control.</summary>
    public int Version { get; }

    /// <summary>HTML reconstructed from <c>htmltag</c> and <c>mhtmltag</c> destinations.</summary>
    public string Html { get; }
}
