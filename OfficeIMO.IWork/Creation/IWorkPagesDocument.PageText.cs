namespace OfficeIMO.IWork;

public sealed partial class IWorkPagesDocument {
    /// <summary>Gets the immutable left header run, or null when headers are disabled.</summary>
    public IWorkTextRunData? Header { get; private set; }
    /// <summary>Gets the immutable left footer run, or null when footers are disabled.</summary>
    public IWorkTextRunData? Footer { get; private set; }
    /// <summary>Gets the native header distance from the page's top edge, in points.</summary>
    public float HeaderDistancePoints { get; private set; }
    /// <summary>Gets the native footer distance from the page's bottom edge, in points.</summary>
    public float FooterDistancePoints { get; private set; }

    /// <summary>Sets one left-aligned header paragraph for every page in the document's single section.</summary>
    /// <remarks>Text must contain at least one character and uses one font and color. Line breaks, tabs, page-number fields and alternate-page templates are not authored.
    /// The distance must be nonnegative and smaller than the top body margin. Native wrapping depends on fonts and text width.</remarks>
    public void SetHeader(string text, string fontName = "Arial", float fontSizePoints = 12,
        string color = "000000", float distancePoints = 36) =>
        SetHeader(new IWorkTextRunData(text, fontName, fontSizePoints, color), distancePoints);

    /// <summary>Sets a snapshot of one immutable character run as the left header for every page.</summary>
    /// <remarks>Supports the run's decorations. The same single-paragraph and distance bounds as the string overload apply.</remarks>
    public void SetHeader(IWorkTextRunData run, float distancePoints = 36) {
        ValidatePageText(run, distancePoints, TopMarginPoints);
        Header = run; HeaderDistancePoints = distancePoints;
    }

    /// <summary>Sets one left-aligned footer paragraph for every page in the document's single section.</summary>
    /// <remarks>Text must contain at least one character and uses one font and color. Line breaks, tabs, page-number fields and alternate-page templates are not authored.
    /// The distance must be nonnegative and smaller than the bottom body margin. Native wrapping depends on fonts and text width.</remarks>
    public void SetFooter(string text, string fontName = "Arial", float fontSizePoints = 12,
        string color = "000000", float distancePoints = 36) =>
        SetFooter(new IWorkTextRunData(text, fontName, fontSizePoints, color), distancePoints);

    /// <summary>Sets a snapshot of one immutable character run as the left footer for every page.</summary>
    /// <remarks>Supports the run's decorations. The same single-paragraph and distance bounds as the string overload apply.</remarks>
    public void SetFooter(IWorkTextRunData run, float distancePoints = 36) {
        ValidatePageText(run, distancePoints, BottomMarginPoints);
        Footer = run; FooterDistancePoints = distancePoints;
    }

    /// <summary>Disables the header and clears its text and distance.</summary>
    public void ClearHeader() { Header = null; HeaderDistancePoints = 0; }
    /// <summary>Disables the footer and clears its text and distance.</summary>
    public void ClearFooter() { Footer = null; FooterDistancePoints = 0; }

    private static void ValidatePageText(IWorkTextRunData run, float distancePoints, float bodyMargin) {
        if (run == null) throw new ArgumentNullException(nameof(run));
        if (run.Text.Length == 0)
            throw new ArgumentException("Header and footer text must contain at least one character; clear the area to disable it.", nameof(run));
        if (run.Text.IndexOfAny(new[] { '\n', '\t', '\u2028', '\u2029' }) >= 0)
            throw new ArgumentException("A header or footer run cannot contain line breaks or tabs.", nameof(run));
        IWorkCreationValidation.ValidateFinite(distancePoints, nameof(distancePoints));
        if (distancePoints < 0 || distancePoints >= bodyMargin)
            throw new ArgumentOutOfRangeException(nameof(distancePoints), "The page-text distance must lie inside its body margin.");
    }
}
