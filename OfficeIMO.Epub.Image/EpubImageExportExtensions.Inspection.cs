using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

public static partial class EpubImageExportExtensions {
    /// <summary>Renders one retained fixed-layout XHTML chapter using its declared numeric viewport and
    /// checks rendered element rectangles against that canvas. Uses the shared HTML/Drawing engines and
    /// package resource resolver. No image is encoded. SVG spine inspection, clipped-content/ink overflow,
    /// individual region overflow and native-reader equivalence are not established by this check.</summary>
    /// <param name="source">Publication loaded with raw HTML and resource payloads retained.</param>
    /// <param name="chapterIndex">Zero-based chapter index.</param>
    /// <param name="options">Font, resource and safety settings. Mode, viewport and margins are set from the page;
    /// chapter selection and image-output settings do not apply. External asynchronous resources remain diagnosed.</param>
    /// <param name="cancellationToken">Cooperative cancellation.</param>
    public static EpubFixedLayoutInspection InspectFixedLayoutPage(this EpubDocument source, int chapterIndex,
        EpubImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        cancellationToken.ThrowIfCancellationRequested();
        if (chapterIndex < 0 || chapterIndex >= source.Chapters.Count) throw new ArgumentOutOfRangeException(nameof(chapterIndex));
        EpubChapter chapter = source.Chapters[chapterIndex];
        if (!chapter.IsFixedLayout || !string.Equals(chapter.MediaType, "application/xhtml+xml", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Fixed-layout inspection requires an XHTML spine chapter declared pre-paginated.");
        if (chapter.Encryption?.RequiresDecryption == true || string.IsNullOrWhiteSpace(chapter.Html))
            throw new NotSupportedException("Fixed-layout inspection requires retained, decrypted chapter XHTML; text fallback cannot establish geometry.");
        EpubImageExportOptions effective = options?.CloneEpub() ?? new EpubImageExportOptions();
        effective.Validate();
        (double width, double height) = ReadInspectionViewport(chapter.Html!, effective.MaxInputCharacters);
        if (width > effective.MaxSurfaceWidth || height > effective.MaxSurfaceHeight)
            throw new InvalidDataException("The declared viewport exceeds the configured render surface limit.");
        effective.Mode = HtmlRenderMode.Continuous;
        effective.ViewportWidth = width; effective.ViewportHeight = height; effective.Margins = HtmlRenderMargins.All(0D);
        EpubChapterRenderPreparation preparation = PrepareChapter(chapter, effective, BuildResourceIndex(source, cancellationToken));
        HtmlRenderDocument rendering = HtmlRenderEngine.Render(preparation.Document, preparation.Options, cancellationToken);
        HtmlRenderPage page = rendering.Pages.Single();
        OfficeDrawingQualityReport quality = page.InspectCanvasBounds(width, height, effective.MaxSurfaceWidth,
            effective.MaxSurfaceHeight, cancellationToken);
        return new EpubFixedLayoutInspection(chapter.Path, width, height, rendering, quality, source.Diagnostics, preparation.Diagnostics);
    }

    private static (double Width, double Height) ReadInspectionViewport(string source, int maximumCharacters) {
        if (source.Length > maximumCharacters) throw new InvalidDataException("Chapter XHTML exceeds the configured character limit.");
        using var input = new StringReader(source);
        using var reader = XmlReader.Create(input, new XmlReaderSettings { DtdProcessing = DtdProcessing.Ignore,
            XmlResolver = null, MaxCharactersInDocument = maximumCharacters });
        XDocument document = XDocument.Load(reader);
        XNamespace html = "http://www.w3.org/1999/xhtml";
        if (document.Root?.Name != html + "html") throw new InvalidDataException("Fixed-layout inspection requires namespaced XHTML.");
        XElement[] viewports = document.Root.Elements(html + "head").Elements(html + "meta")
            .Where(e => string.Equals((string?)e.Attribute("name"), "viewport", StringComparison.OrdinalIgnoreCase)).ToArray();
        if (viewports.Length != 1) throw new InvalidDataException("Fixed-layout inspection requires exactly one viewport declaration.");
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (string part in ((string?)viewports[0].Attribute("content") ?? string.Empty).Split(new[] { ',', ';' })) {
            string[] pair = part.Split('=');
            if (pair.Length != 2) throw new InvalidDataException("Viewport declarations require unambiguous key=value pairs.");
            string key = pair[0].Trim();
            if (values.ContainsKey(key)) throw new InvalidDataException("Viewport declarations cannot repeat a property.");
            values.Add(key, pair[1].Trim());
        }
        double Dimension(string name) {
            if (!values.TryGetValue(name, out string? value) ||
                !double.TryParse(value, NumberStyles.AllowDecimalPoint, CultureInfo.InvariantCulture, out double result) ||
                double.IsNaN(result) || double.IsInfinity(result) || result <= 0D)
                throw new InvalidDataException("Fixed-layout viewport " + name + " must be a positive numeric CSS-pixel dimension.");
            return result;
        }
        return (Dimension("width"), Dimension("height"));
    }
}
