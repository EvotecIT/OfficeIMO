using OfficeIMO.ContentSafety;
using System.Threading;

namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    /// <summary>Inspects the loaded native package without serialization, external-resource access, or mutation. Source-load limits remain the caller's responsibility.</summary>
    public OfficeContentSafetyReport InspectLoadedContentSafety(OfficeContentSafetyOptions? options = null, CancellationToken cancellationToken = default) =>
        InspectOdfContentSafety(this, options, targets: null, cancellationToken);
}
