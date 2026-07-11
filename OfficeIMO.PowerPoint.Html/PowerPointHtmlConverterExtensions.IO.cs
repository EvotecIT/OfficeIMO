using OfficeIMO.Html;
using System.Threading;
using System.Threading.Tasks;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class PowerPointHtmlConverterExtensions {
    /// <summary>Writes a presentation as UTF-8 HTML to a stream and leaves it open.</summary>
    public static void SaveAsHtml(this PptCore.PowerPointPresentation presentation, Stream stream, PowerPointHtmlSaveOptions? options = null) {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        HtmlTextIO.Write(stream, presentation.ToHtml(options));
    }

    /// <summary>Asynchronously writes a presentation as UTF-8 HTML.</summary>
    public static async Task SaveAsHtmlAsync(this PptCore.PowerPointPresentation presentation, string path, PowerPointHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        cancellationToken.ThrowIfCancellationRequested();
        await HtmlTextIO.WriteAsync(path, presentation.ToHtml(options), cancellationToken).ConfigureAwait(false);
    }

    /// <summary>Asynchronously writes a presentation as UTF-8 HTML to a stream and leaves it open.</summary>
    public static async Task SaveAsHtmlAsync(this PptCore.PowerPointPresentation presentation, Stream stream, PowerPointHtmlSaveOptions? options = null, CancellationToken cancellationToken = default) {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        cancellationToken.ThrowIfCancellationRequested();
        await HtmlTextIO.WriteAsync(stream, presentation.ToHtml(options), cancellationToken).ConfigureAwait(false);
    }
}
