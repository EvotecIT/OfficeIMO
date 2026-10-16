using PdfCore = OfficeIMO.Pdf;
using System.Threading;

namespace OfficeIMO.Html.Pdf;

/// <summary>PDF output paired with the exact retained render request result that produced it.</summary>
public sealed class HtmlPdfRenderRequestResult {
    internal HtmlPdfRenderRequestResult(HtmlRenderResult renderResult, PdfCore.PdfDocumentConversionResult output) {
        RenderResult = renderResult ?? throw new ArgumentNullException(nameof(renderResult));
        Output = output ?? throw new ArgumentNullException(nameof(output));
    }

    /// <summary>Resolved request, selected surfaces, source offsets, clipping, and HTML diagnostics.</summary>
    public HtmlRenderResult RenderResult { get; }

    /// <summary>PDF document and combined HTML/PDF conversion report.</summary>
    public PdfCore.PdfDocumentConversionResult Output { get; }

    /// <summary>Serializes the completed PDF document.</summary>
    public byte[] ToBytes(CancellationToken cancellationToken = default) => Output.ToBytes(cancellationToken);
}
