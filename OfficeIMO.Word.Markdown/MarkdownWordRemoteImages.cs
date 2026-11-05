using OfficeIMO.Drawing;
using OfficeIMO.Markdown;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Word.Markdown;

internal static class MarkdownWordRemoteImages {
    internal static async Task PrepareAsync(MarkdownDocument document, MarkdownWordExportOptions options, CancellationToken token) {
        if (!options.AllowRemoteImages || options.RemoteImageResolver != null) return;
        var collector = new Collector(options);
        collector.Visit(document);
        var downloads = new Dictionary<Uri, (byte[]? Bytes, Exception? Error)>();
        foreach (Uri uri in collector.Images) {
            token.ThrowIfCancellationRequested();
            if (!options.AllowedImageSchemes.Contains(uri.Scheme) || options.ImageUrlValidator?.Invoke(uri) == false) continue;
            try {
                OfficeRemoteImageLoadOptions reading = MarkdownWordExportOptions.CopyRemoteImageOptions(options.RemoteImageLoadOptions) ?? new();
                if (options.MaximumRemoteImageBytes >= 0) reading.MaximumBytes = Math.Min(reading.MaximumBytes, options.MaximumRemoteImageBytes);
                OfficeRemoteImage image = await OfficeRemoteImageLoader.LoadAsync(uri, reading, token).ConfigureAwait(false);
                downloads.Add(uri, (image.ToBytes(), null));
            } catch (Exception error) when (error is not OperationCanceledException) { downloads.Add(uri, (null, error)); }
        }
        token.ThrowIfCancellationRequested();
        // This operation's resolver contains completed downloads only, never a live document or asynchronous I/O.
        options.RemoteImageResolver = uri => {
            if (!downloads.TryGetValue(uri, out var image)) return null;
            if (image.Error != null) throw new InvalidOperationException(image.Error.Message, image.Error);
            return image.Bytes;
        };
    }

    private sealed class Collector(MarkdownWordExportOptions options) : MarkdownVisitor {
        internal HashSet<Uri> Images { get; } = new();
        protected override void VisitImageBlock(ImageBlock block) { Add(block.Path); base.VisitImageBlock(block); }
        protected override void VisitImageInline(ImageInline inline) { Add(inline.Src); base.VisitImageInline(inline); }
        protected override void VisitImageLinkInline(ImageLinkInline inline) { Add(inline.ImageUrl); base.VisitImageLinkInline(inline); }
        private void Add(string value) {
            Uri? uri;
            if (!Uri.TryCreate(value, UriKind.Absolute, out uri) && Uri.TryCreate(options.BaseUri, UriKind.Absolute, out var baseUri))
                Uri.TryCreate(baseUri, value, out uri);
            if (uri != null && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps)) Images.Add(uri);
        }
    }
}
