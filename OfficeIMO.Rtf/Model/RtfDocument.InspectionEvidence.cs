using OfficeIMO.ContentSafety;
using System.Threading;

namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    /// <summary>Inspects the loaded semantic document without reparsing, saving or executing object content. Source-load limits remain the caller's responsibility.</summary>
    public OfficeContentSafetyReport InspectLoadedContentSafety(OfficeContentSafetyOptions? options = null, CancellationToken cancellationToken = default) =>
        InspectRtfContentSafety(this, options, targets: null, cancellationToken);
}
