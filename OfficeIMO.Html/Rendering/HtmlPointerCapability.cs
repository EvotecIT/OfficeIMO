namespace OfficeIMO.Html;

/// <summary>Primary pointing-device accuracy exposed to CSS media queries.</summary>
public enum HtmlPointerCapability {
    /// <summary>No pointing device is available to the static output.</summary>
    None,
    /// <summary>The primary pointing device has limited accuracy.</summary>
    Coarse,
    /// <summary>The primary pointing device has fine accuracy.</summary>
    Fine
}
