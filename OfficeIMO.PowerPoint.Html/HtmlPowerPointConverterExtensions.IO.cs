using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class HtmlPowerPointConverterExtensions {
    /// <summary>Reads semantic HTML from a stream and imports a presentation.</summary>
    public static PptCore.PowerPointPresentation ToPowerPointPresentation(this Stream htmlStream, HtmlToPowerPointOptions? options = null) =>
        HtmlTextIO.Read(htmlStream).ToPowerPointPresentation(options);

    /// <summary>Reads semantic HTML from a stream and returns a presentation plus structured evidence.</summary>
    public static HtmlToPowerPointResult ToPowerPointPresentationResult(this Stream htmlStream, HtmlToPowerPointOptions? options = null) =>
        HtmlTextIO.Read(htmlStream).ToPowerPointPresentationResult(options);

    /// <summary>Asynchronously reads semantic HTML from a stream and imports a presentation.</summary>
    public static async Task<PptCore.PowerPointPresentation> ToPowerPointPresentationAsync(this Stream htmlStream, HtmlToPowerPointOptions? options = null, CancellationToken cancellationToken = default) =>
        (await HtmlTextIO.ReadAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToPowerPointPresentation(options);

    /// <summary>Asynchronously reads semantic HTML from a stream and returns structured evidence.</summary>
    public static async Task<HtmlToPowerPointResult> ToPowerPointPresentationResultAsync(this Stream htmlStream, HtmlToPowerPointOptions? options = null, CancellationToken cancellationToken = default) =>
        (await HtmlTextIO.ReadAsync(htmlStream, cancellationToken).ConfigureAwait(false)).ToPowerPointPresentationResult(options);
}
