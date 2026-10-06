using System.Globalization;
using System.Xml;
using System.Xml.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Epub.Image;

public static partial class EpubImageExportExtensions {
    /// <summary>Renders one retained fixed-layout XHTML chapter using its declared numeric viewport and
    /// checks rendered element rectangles against that canvas. Uses the shared HTML/Drawing engines and
    /// package resource resolver. No image is encoded. Rectangular scene clipping is reported separately.
    /// Positioned XHTML text ink is reported separately. SVG spine inspection, precise path-clipping geometry, individual region overflow
    /// and native-reader equivalence are not established by this check.</summary>
    /// <param name="source">Publication loaded with raw HTML and resource payloads retained.</param>
    /// <param name="chapterIndex">Zero-based chapter index.</param>
    /// <param name="options">Font, resource and safety settings. Mode, viewport and margins are set from the page;
    /// chapter selection and image-output settings do not apply. External asynchronous resources remain diagnosed.</param>
    /// <param name="cancellationToken">Cooperative cancellation.</param>
    public static EpubFixedLayoutInspection InspectFixedLayoutPage(this EpubDocument source, int chapterIndex,
        EpubImageExportOptions? options = null, CancellationToken cancellationToken = default) =>
        InspectFixedLayout(source, chapterIndex, Array.Empty<string>(), options, cancellationToken);

    /// <summary>Inspects the page canvas and selected identified positioned, floating, flex or grid regions
    /// in one render pass. Region findings use local border-box coordinates, before the region's own and
    /// ancestor transforms. Descendant transforms and authored clips are retained. Ancestor clips do not
    /// redefine local containment. Positioned XHTML ink is reported separately, including rectangular/convex contour clipping; unsupported paths remain unmeasured.
    /// Missing, duplicate, hidden or unsupported region targets fail explicitly.</summary>
    /// <param name="source">Publication loaded with raw HTML and resource payloads retained.</param>
    /// <param name="chapterIndex">Zero-based fixed-layout XHTML chapter index.</param>
    /// <param name="regionElementIds">One to 1024 distinct HTML element IDs to inspect.</param>
    /// <param name="options">Font, resource and safety settings, as for page inspection.</param>
    /// <param name="cancellationToken">Cooperative cancellation.</param>
    public static EpubFixedLayoutInspection InspectFixedLayoutRegions(this EpubDocument source, int chapterIndex,
        IReadOnlyList<string> regionElementIds, EpubImageExportOptions? options = null, CancellationToken cancellationToken = default) {
        if (regionElementIds == null) throw new ArgumentNullException(nameof(regionElementIds));
        if (regionElementIds.Count == 0 || regionElementIds.Count > 1024) throw new ArgumentOutOfRangeException(nameof(regionElementIds));
        string[] ids = regionElementIds.ToArray();
        if (ids.Any(id => string.IsNullOrWhiteSpace(id) || id.Length > 1024) || ids.Distinct(StringComparer.Ordinal).Count() != ids.Length)
            throw new ArgumentException("Region IDs must be nonempty, distinct and at most 1024 characters each.", nameof(regionElementIds));
        return InspectFixedLayout(source, chapterIndex, ids, options, cancellationToken);
    }

    private static EpubFixedLayoutInspection InspectFixedLayout(EpubDocument source, int chapterIndex,
        IReadOnlyList<string> regionIds, EpubImageExportOptions? options, CancellationToken cancellationToken) {
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
        EpubChapterRenderPreparation preparation = PrepareChapter(chapter, effective, BuildResourceIndex(source, cancellationToken), cancellationToken);
        HtmlRenderDocument rendering = HtmlRenderEngine.RenderForRegionInspection(preparation.Document, preparation.Options, regionIds, cancellationToken);
        HtmlRenderPage page = rendering.Pages.Single();
        OfficeDrawingQualityReport quality = page.InspectCanvasBounds(width, height, effective.MaxSurfaceWidth,
            effective.MaxSurfaceHeight, cancellationToken);
        var regions = regionIds.Select(id => {
            HtmlRenderPage local = page.CreateRegionInspectionPage(id, cancellationToken);
            return new EpubFixedLayoutRegionInspection(id,
                local.InspectCanvasBounds(local.Width, local.Height, effective.MaxSurfaceWidth, effective.MaxSurfaceHeight, cancellationToken),
                local.InspectTextInk(local.Width, local.Height, cancellationToken, effective.TextShapingProvider,
                    effective.TextShapingLanguage, isRegion: true));
        }).ToArray();
        return new EpubFixedLayoutInspection(chapter.Path, width, height, rendering, quality, source.Diagnostics, preparation.Diagnostics, regions,
            page.InspectClipping(effective.MaxSurfaceWidth, effective.MaxSurfaceHeight, cancellationToken),
            page.InspectTextInk(width, height, cancellationToken, effective.TextShapingProvider, effective.TextShapingLanguage));
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
