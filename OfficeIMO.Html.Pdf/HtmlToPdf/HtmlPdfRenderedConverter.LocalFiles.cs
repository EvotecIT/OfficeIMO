using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    // A named delegate target distinguishes a prepared implicit file loader from a
    // configured host or package resolver when a public render request is reused.
    private sealed class LocalFileResourceLoader {
        internal LocalFileResourceLoader(HtmlRenderSynchronousResourceResolver? packageResolver,
            HtmlRenderResourceResolver? configuredResolver, long maximumBytes) {
            PackageResolver = packageResolver;
            _configuredResolver = configuredResolver;
            _maximumBytes = maximumBytes;
        }

        internal HtmlRenderSynchronousResourceResolver? PackageResolver { get; }
        private readonly HtmlRenderResourceResolver? _configuredResolver;
        private readonly long _maximumBytes;

        internal bool ResolveSynchronously(HtmlRenderResourceRequest request, CancellationToken token,
            out HtmlResolvedResource? resource) {
            if (PackageResolver != null && PackageResolver(request, token, out resource)) return true;
            return TryReadLocalFileResource(request, _maximumBytes, token, out resource);
        }

        internal async Task<HtmlResolvedResource?> ResolveAsync(HtmlRenderResourceRequest request, CancellationToken token) {
            if (_configuredResolver != null) {
                HtmlResolvedResource? configured = await _configuredResolver(request, token).ConfigureAwait(false);
                if (configured != null) return configured;
            }
            return TryReadLocalFileResource(request, _maximumBytes, token, out HtmlResolvedResource? local) ? local : null;
        }
    }
}
