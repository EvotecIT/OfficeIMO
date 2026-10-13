using OfficeIMO.ContentSafety;
using System.Threading;

namespace OfficeIMO.Excel;

public partial class ExcelDocument {
    /// <summary>Structural observations captured from the exact loader source bytes when requested; subsequent edits do not change this source report.</summary>
    public OfficePackageSecurityReport? SourcePackageSecurityReport { get; private set; }
    /// <summary>Identifies the physical source scope of captured package observations.</summary>
    public string? SourcePackageSecurityScope { get; private set; }

    private static ExcelDocument WithSourcePackageInspection(ExcelDocument document, byte[] sourceBytes, ExcelLoadOptions options, string scope = "source-container") {
        if (!options.CaptureSourcePackageSecurity) return document;
        try {
            document.SourcePackageSecurityReport = OfficePackageSecurityInspector.Inspect(sourceBytes, options.PackageSecurity);
            document.SourcePackageSecurityScope = scope;
            return document;
        } catch { document.Dispose(); throw; }
    }
    /// <summary>Inspects the currently loaded workbook without reopening, calculating formulas, refreshing links, or saving. Source-load limits remain the caller's responsibility.</summary>
    public OfficeContentSafetyReport InspectLoadedContentSafety(OfficeContentSafetyOptions? options = null, CancellationToken cancellationToken = default) =>
        InspectContentSafetyDocument(this, options, targets: null, cancellationToken);
}
