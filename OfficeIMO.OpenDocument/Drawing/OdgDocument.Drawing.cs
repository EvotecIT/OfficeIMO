using OfficeIMO.Drawing;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgDocument {
    /// <summary>Projects every Draw page in document order with page-qualified feature losses.</summary>
    /// <param name="lossPolicy">Treatment of unsupported content and approximations.</param>
    /// <param name="forPrint">Selects print-visible layers instead of screen-visible layers.</param>
    /// <param name="maximumPages">Maximum source page count accepted before projecting any page.</param>
    /// <param name="cancellationToken">Cooperative cancellation between pages and shapes.</param>
    public OdfConversionResult<IReadOnlyList<OfficeDrawing>> ToDrawings(
        OdfConversionLossPolicy lossPolicy = OdfConversionLossPolicy.ReportOnly, bool forPrint = false,
        int maximumPages = 1000, CancellationToken cancellationToken = default) =>
        ToDrawings(lossPolicy, forPrint, maximumPages, cancellationToken, null);

    /// <summary>Projects all pages using one explicit date/time field settings snapshot.</summary>
    public OdfConversionResult<IReadOnlyList<OfficeDrawing>> ToDrawings(OdfConversionLossPolicy lossPolicy,
        bool forPrint, int maximumPages, CancellationToken cancellationToken, OdfDateTimeFieldProjectionOptions? dateTimeFields) =>
        ToDrawings(lossPolicy, forPrint, maximumPages, cancellationToken, dateTimeFields, null, null);

    /// <summary>Projects pages using caller-supplied fonts and shaping for both layout decisions and the returned drawings.</summary>
    /// <remarks>The immutable profile owns a font snapshot. Later changes to a returned drawing's fonts or shaping can change its layout without updating this operation's loss report.</remarks>
    public OdfConversionResult<IReadOnlyList<OfficeDrawing>> ToDrawings(OfficeRenderingProfile renderingProfile,
        OdfConversionLossPolicy lossPolicy = OdfConversionLossPolicy.ReportOnly, bool forPrint = false,
        int maximumPages = 1000, CancellationToken cancellationToken = default, OdfDateTimeFieldProjectionOptions? dateTimeFields = null) {
        if (renderingProfile == null) throw new ArgumentNullException(nameof(renderingProfile));
        return ToDrawings(lossPolicy, forPrint, maximumPages, cancellationToken, dateTimeFields, renderingProfile, null);
    }

    // Format bridges can supply the authoritative metrics of their final renderer without exposing callbacks publicly.
    internal OdfConversionResult<IReadOnlyList<OfficeDrawing>> ToDrawings(OdfConversionLossPolicy lossPolicy,
        bool forPrint, int maximumPages, CancellationToken cancellationToken, OdfDateTimeFieldProjectionOptions? dateTimeFields,
        OfficeRenderingProfile? renderingProfile, OfficeDrawingTextMetrics? layoutMetrics) {
        cancellationToken.ThrowIfCancellationRequested();
        if (maximumPages < 1) throw new ArgumentOutOfRangeException(nameof(maximumPages));
        if (lossPolicy < OdfConversionLossPolicy.ReportOnly || lossPolicy > OdfConversionLossPolicy.ThrowOnAnyLoss)
            throw new ArgumentOutOfRangeException(nameof(lossPolicy));
        IReadOnlyList<OdgPage> pages = Pages;
        if (pages.Count == 0) throw new InvalidOperationException("Draw page projection requires at least one source page.");
        if (pages.Count > maximumPages)
            throw new InvalidOperationException("The drawing exceeds the configured " + maximumPages.ToString(CultureInfo.InvariantCulture) + "-page projection limit.");
        var fieldProjection = new OdfDateTimeFieldProjection(this, dateTimeFields);
        var report = new OdfConversionReport("ODG", "OfficeDrawing");
        if (GetXml("meta.xml").Root?.Element(OdfNamespaces.Office + "meta")?.Elements()
            .Any(element => element.Name != OdfNamespaces.Meta + "generator" && element.Name != OdfNamespaces.Meta + "document-statistic") == true)
            report.Add("document-metadata", OdfConversionMappingStatus.Skipped,
                message: "Source metadata is retained in ODF but is not copied into the drawing-page projection.");
        var drawings = new OfficeDrawing[pages.Count];
        for (int index = 0; index < pages.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            OdgPage page = pages[index];
            var projected = page.ToDrawing(OdfConversionLossPolicy.ReportOnly, forPrint, cancellationToken, index + 1, pages.Count, fieldProjection, renderingProfile, layoutMetrics);
            string location = "page:" + (index + 1).ToString(CultureInfo.InvariantCulture) + ":" + page.Name;
            foreach (OdfConversionMapping mapping in projected.Report.Mappings)
                report.Add(location + "/" + mapping.Feature, mapping.Status, mapping.Count, mapping.Message);
            report.Add(location, OdfConversionMappingStatus.Converted, message: "The source page is projected at its original dimensions using " + (forPrint ? "print" : "screen") + " layer visibility.");
            report.ApplyPolicy(lossPolicy);
            drawings[index] = projected.Value;
        }
        cancellationToken.ThrowIfCancellationRequested();
        return new OdfConversionResult<IReadOnlyList<OfficeDrawing>>(Array.AsReadOnly(drawings), report);
    }
}
