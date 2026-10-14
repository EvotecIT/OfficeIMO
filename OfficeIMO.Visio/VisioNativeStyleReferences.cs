using System.Xml;
using System.Xml.Linq;

namespace OfficeIMO.Visio;

/// <summary>
/// Retains native style references, including omitted attributes that inherit from a master or document.
/// A null source identifies a newly authored element and permits save-time defaults.
/// </summary>
internal sealed class VisioNativeStyleReferences {
    private VisioNativeStyleReferences(string? lineStyle, string? fillStyle, string? textStyle) {
        LineStyle = lineStyle;
        FillStyle = fillStyle;
        TextStyle = textStyle;
    }

    internal string? LineStyle { get; }
    internal string? FillStyle { get; }
    internal string? TextStyle { get; }

    /// <summary>Captures a loaded header even when all style attributes are absent.</summary>
    internal static VisioNativeStyleReferences Read(XElement element) => new(
        (string?)element.Attribute("LineStyle"),
        (string?)element.Attribute("FillStyle"),
        (string?)element.Attribute("TextStyle"));

    /// <summary>Writes native references, applying authored defaults only when no native header was loaded.</summary>
    internal static void Write(XmlWriter writer, VisioNativeStyleReferences? source,
        string? lineStyle = "0", string? fillStyle = "0", string? textStyle = "0") {
        string? line = source == null ? lineStyle : source.LineStyle;
        string? fill = source == null ? fillStyle : source.FillStyle;
        string? text = source == null ? textStyle : source.TextStyle;
        if (line != null) writer.WriteAttributeString("LineStyle", line);
        if (fill != null) writer.WriteAttributeString("FillStyle", fill);
        if (text != null) writer.WriteAttributeString("TextStyle", text);
    }
}
