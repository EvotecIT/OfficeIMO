using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

namespace PSWriteOffice.Services.Markdown;

internal static class MarkdownRemoteImageService
{
    private static readonly HttpClient Client = new();

    internal static async Task ConfigureResolverAsync(
        MarkdownDoc document,
        MarkdownToWordOptions options,
        CancellationToken cancellationToken)
    {
        var collector = new RemoteImageCollector();
        collector.Visit(document);

        var images = collector.Images
            .Where(uri => options.AllowedImageSchemes.Contains(uri.Scheme))
            .Where(uri => options.ImageUrlValidator == null || options.ImageUrlValidator(uri))
            .Distinct()
            .ToArray();

        if (images.Length == 0)
        {
            return;
        }

        var downloads = new Dictionary<Uri, RemoteImageDownload>();
        foreach (var image in images)
        {
            cancellationToken.ThrowIfCancellationRequested();
            try
            {
                downloads[image] = RemoteImageDownload.Success(
                    await DownloadAsync(image, options.MaximumRemoteImageBytes, cancellationToken).ConfigureAwait(false));
            }
            catch (Exception ex) when (!(ex is OperationCanceledException))
            {
                downloads[image] = RemoteImageDownload.Failure(ex.Message);
            }
        }

        options.RemoteImageResolver = uri =>
        {
            if (!downloads.TryGetValue(uri, out var download))
            {
                return null;
            }

            if (download.Error != null)
            {
                throw new InvalidOperationException(download.Error);
            }

            return download.Bytes;
        };
    }

    private static async Task<byte[]> DownloadAsync(Uri uri, long maximumBytes, CancellationToken cancellationToken)
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, uri);
        using var response = await Client.SendAsync(request, HttpCompletionOption.ResponseHeadersRead, cancellationToken).ConfigureAwait(false);
        response.EnsureSuccessStatusCode();

        var contentLength = response.Content.Headers.ContentLength;
        if (maximumBytes >= 0 && contentLength.HasValue && contentLength.Value > maximumBytes)
        {
            throw new InvalidOperationException($"The remote image is {contentLength.Value} bytes, exceeding the configured limit of {maximumBytes} bytes.");
        }

        using var source = await response.Content.ReadAsStreamAsync().ConfigureAwait(false);
        using var destination = new MemoryStream();
        var buffer = new byte[81920];
        long total = 0;
        int read;
        while ((read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false)) > 0)
        {
            total += read;
            if (maximumBytes >= 0 && total > maximumBytes)
            {
                throw new InvalidOperationException($"The remote image exceeds the configured limit of {maximumBytes} bytes.");
            }

            await destination.WriteAsync(buffer, 0, read, cancellationToken).ConfigureAwait(false);
        }

        return destination.ToArray();
    }

    private sealed class RemoteImageCollector : MarkdownVisitor
    {
        internal List<Uri> Images { get; } = new();

        protected override void VisitImageBlock(ImageBlock block)
        {
            Add(block.Path);
            base.VisitImageBlock(block);
        }

        protected override void VisitImageInline(ImageInline inline)
        {
            Add(inline.Src);
            base.VisitImageInline(inline);
        }

        protected override void VisitImageLinkInline(ImageLinkInline inline)
        {
            Add(inline.ImageUrl);
            base.VisitImageLinkInline(inline);
        }

        private void Add(string value)
        {
            if (Uri.TryCreate(value, UriKind.Absolute, out var uri)
                && (uri.Scheme == Uri.UriSchemeHttp || uri.Scheme == Uri.UriSchemeHttps))
            {
                Images.Add(uri);
            }
        }
    }

    private sealed class RemoteImageDownload
    {
        private RemoteImageDownload(byte[]? bytes, string? error)
        {
            Bytes = bytes;
            Error = error;
        }

        internal byte[]? Bytes { get; }

        internal string? Error { get; }

        internal static RemoteImageDownload Success(byte[] bytes) => new(bytes, null);

        internal static RemoteImageDownload Failure(string error) => new(null, error);
    }
}
