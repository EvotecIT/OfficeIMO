using OfficeIMO.PowerPoint;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services.PowerPoint;

/// <summary>Identifies the native body capturing one footer, so fields cannot leak into other text scopes.</summary>
internal sealed class PowerPointFooterScope {
    internal PowerPointFooterScope(PowerPointDocument document) => Document = document;
    internal PowerPointDocument Document { get; }
    internal static PowerPointFooterScope? Current {
        get {
            var scope = OfficeDocumentContext.Find<PowerPointFooterScope>();
            return scope != null && ReferenceEquals(scope.Document, OfficeDocumentContext.Find<PowerPointDocument>()) ? scope : null;
        }
    }
}
