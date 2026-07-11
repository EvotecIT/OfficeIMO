namespace OfficeIMO.Html;

/// <summary>
/// A source heading retained by the shared render model for navigation-capable backends.
/// </summary>
public sealed class HtmlRenderHeading {
    internal HtmlRenderHeading(int semanticNodeId, int level, string text, int pageNumber, double x, double y) {
        SemanticNodeId = semanticNodeId;
        Level = level;
        Text = text;
        PageNumber = pageNumber;
        X = x;
        Y = y;
    }

    /// <summary>Stable operation-scoped identifier of the source semantic element.</summary>
    public int SemanticNodeId { get; }

    /// <summary>Heading level from 1 through 6.</summary>
    public int Level { get; }

    /// <summary>Rendered heading text, including text split across styled spans or lines.</summary>
    public string Text { get; }

    /// <summary>One-based destination page number.</summary>
    public int PageNumber { get; }

    /// <summary>Left destination coordinate in CSS pixels.</summary>
    public double X { get; }

    /// <summary>Top destination coordinate in CSS pixels.</summary>
    public double Y { get; }

    internal static bool TryGetLevel(string? semanticRole, out int level) {
        switch (semanticRole) {
            case "heading-1": level = 1; return true;
            case "heading-2": level = 2; return true;
            case "heading-3": level = 3; return true;
            case "heading-4": level = 4; return true;
            case "heading-5": level = 5; return true;
            case "heading-6": level = 6; return true;
            default: level = 0; return false;
        }
    }
}
