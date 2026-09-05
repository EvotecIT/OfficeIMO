namespace OfficeIMO.Drawing;

/// <summary>
/// Describes one bounded inline XHTML viewport requested by an SVG <c>foreignObject</c> element.
/// </summary>
public sealed class OfficeSvgForeignObjectContext {
    internal OfficeSvgForeignObjectContext(string html, double width, double height) {
        Html = html;
        Width = width;
        Height = height;
    }

    /// <summary>Serialized child markup contained by the SVG <c>foreignObject</c>.</summary>
    public string Html { get; }

    /// <summary>Requested local viewport width in SVG user units.</summary>
    public double Width { get; }

    /// <summary>Requested local viewport height in SVG user units.</summary>
    public double Height { get; }
}

/// <summary>
/// Renders bounded inline XHTML for an SVG <c>foreignObject</c> into the shared drawing scene.
/// Return <see langword="null"/> when the content cannot be represented.
/// </summary>
/// <param name="context">Inline markup and exact local viewport requested by the SVG reader.</param>
/// <returns>A drawing whose dimensions match the requested viewport, or <see langword="null"/>.</returns>
public delegate OfficeDrawing? OfficeSvgForeignObjectRenderer(OfficeSvgForeignObjectContext context);
