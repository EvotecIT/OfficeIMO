using OfficeIMO.Core.Internal.OpenXml;

namespace OfficeIMO.OpenXml.Internal {
    /// <summary>
    /// A chart part as the shared chart reader and writer see it, wherever the host holds it: an engine's own package
    /// part or an Open XML SDK part. Elements come back as <see cref="OfficeXmlNode"/>, so the shared code reads and edits
    /// both the same way.
    /// </summary>
    internal interface IOfficeChartPart {
        /// <summary>The part's <c>c:chartSpace</c> root element, or none when the part has no content.</summary>
        OfficeXmlNode ChartSpace { get; }

        /// <summary>The roots of the chart style (<c>cs:chartStyle</c>) parts the chart relates to, at most <paramref name="limit"/>.</summary>
        IReadOnlyList<OfficeXmlNode> ChartStyles(int limit);

        /// <summary>
        /// The color map of the slide, layout, master, notes or document that holds the chart, which scheme colors in
        /// the chart go through when the chart space has no color map override of its own. None when the host has none.
        /// </summary>
        OfficeXmlNode HostColorMap { get; }

        /// <summary>The font scheme of the theme in effect for the part that holds the chart, or none.</summary>
        OfficeXmlNode HostFontScheme { get; }

        /// <summary>The embedded packages the chart relates to, in relationship order, at most <paramref name="limit"/>.</summary>
        IReadOnlyList<IOfficeChartEmbeddedPackage> EmbeddedPackages(int limit);

        /// <summary>Records an edit of the chart space, so the host writes it.</summary>
        void MarkEdited();
    }

    /// <summary>A package embedded in a chart part, such as the workbook behind the chart's data.</summary>
    internal interface IOfficeChartEmbeddedPackage {
        /// <summary>The id of the chart part's relationship to the package.</summary>
        string RelationshipId { get; }

        string ContentType { get; }

        /// <summary>Opens the package bytes for reading.</summary>
        Stream OpenRead();

        /// <summary>Replaces the package bytes.</summary>
        void Replace(byte[] data);
    }
}
