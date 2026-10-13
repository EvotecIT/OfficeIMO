using OfficeIMO.ContentSafety;
using System.Threading;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Structural observations captured from the exact loader source bytes when requested; subsequent edits do not change this source report.</summary>
    public OfficePackageSecurityReport? SourcePackageSecurityReport { get; private set; }
    /// <summary>Identifies whether captured observations describe the original source container or its decrypted package.</summary>
    public string? SourcePackageSecurityScope { get; private set; }

    private static WordDocument WithSourcePackageInspection(WordDocument document, byte[] sourceBytes, WordLoadOptions options, string scope = "source-container") {
        if (!options.CaptureSourcePackageSecurity) return document;
        try {
            document.SourcePackageSecurityReport = OfficePackageSecurityInspector.Inspect(sourceBytes, options.PackageSecurity);
            document.SourcePackageSecurityScope = scope;
            return document;
        } catch { document.Dispose(); throw; }
    }
    /// <summary>Inspects the currently loaded document without reopening, saving, or executing its content. Source-load limits remain the caller's responsibility.</summary>
    public OfficeContentSafetyReport InspectLoadedContentSafety(OfficeContentSafetyOptions? options = null, CancellationToken cancellationToken = default) =>
        InspectContentSafetyDocument(_wordprocessingDocument, options, targets: null, cancellationToken);
}
