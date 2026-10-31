using OfficeIMO.ContentSafety;
using System.Threading;

namespace OfficeIMO.PowerPoint;

public sealed partial class PowerPointDocument {
    /// <summary>Structural observations captured from the exact loader source bytes when requested; subsequent edits do not change this source report.</summary>
    public OfficePackageSecurityReport? SourcePackageSecurityReport { get; private set; }
    /// <summary>Identifies the physical source scope of captured package observations.</summary>
    public string? SourcePackageSecurityScope { get; private set; }

    private static PowerPointDocument WithSourcePackageInspection(PowerPointDocument presentation, byte[] sourceBytes, PowerPointLoadOptions options, string scope = "source-container") {
        if (!options.CaptureSourcePackageSecurity) return presentation;
        try {
            presentation.SourcePackageSecurityReport = OfficePackageSecurityInspector.Inspect(sourceBytes, options.PackageSecurity);
            presentation.SourcePackageSecurityScope = scope;
            return presentation;
        } catch { presentation.Dispose(); throw; }
    }
    /// <summary>Inspects the currently loaded presentation without reopening, saving or running content. Source-load limits remain the caller's responsibility.</summary>
    public OfficeContentSafetyReport InspectLoadedContentSafety(OfficeContentSafetyOptions? options = null, CancellationToken cancellationToken = default) =>
        InspectContentSafetyDocument(this, options, targets: null, cancellationToken);
}
