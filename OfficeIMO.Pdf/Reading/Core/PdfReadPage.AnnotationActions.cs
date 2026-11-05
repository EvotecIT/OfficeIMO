using System.Threading;

namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private string? ReadAnnotationUri(PdfDictionary? action, CancellationToken cancellationToken) {
        if (action?.Get<PdfName>("S")?.Name != "URI") return null;
        return TryGetString(action.Items.TryGetValue("URI", out var target) ? target : null, out string? uri)
            && Guard.IsUriAction(uri, cancellationToken) ? uri : null;
    }
}
