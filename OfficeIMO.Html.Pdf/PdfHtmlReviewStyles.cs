namespace OfficeIMO.Html.Pdf;

/// <summary>Loads the adapter-owned stylesheet layered over the shared OfficeIMO HTML shell.</summary>
internal static class PdfHtmlReviewStyles {
    private const string PositioningResourceName = "OfficeIMO.Html.Pdf.Assets.PdfHtmlPositioning.css";
    private const string ReviewResourceName = "OfficeIMO.Html.Pdf.Assets.PdfHtmlReview.css";
    private static readonly Lazy<string> PositioningStyles = Create(PositioningResourceName);
    private static readonly Lazy<string> ReviewStyles = Create(ReviewResourceName);

    internal static string GetPositioning(string newLine = "\n") => Normalize(PositioningStyles.Value, newLine);

    internal static string GetReview(string newLine = "\n") => Normalize(ReviewStyles.Value, newLine);

    private static Lazy<string> Create(string resourceName) => new Lazy<string>(
        () => Load(resourceName),
        System.Threading.LazyThreadSafetyMode.ExecutionAndPublication);

    private static string Normalize(string value, string newLine) {
        string resolvedNewLine = string.IsNullOrEmpty(newLine) ? "\n" : newLine;
        return value
            .Replace("\r\n", "\n")
            .Replace('\r', '\n')
            .Replace("\n", resolvedNewLine);
    }

    private static string Load(string resourceName) {
        using Stream stream = typeof(PdfHtmlReviewStyles).Assembly.GetManifestResourceStream(resourceName)
            ?? throw new InvalidOperationException("Embedded PDF-to-HTML stylesheet is missing: " + resourceName + ".");
        using var reader = new StreamReader(stream, Encoding.UTF8, true);
        return reader.ReadToEnd();
    }
}
