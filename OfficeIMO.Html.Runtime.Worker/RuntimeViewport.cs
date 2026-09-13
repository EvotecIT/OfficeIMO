using AngleSharp.Dom;
using AngleSharp.Html.Dom;

namespace OfficeIMO.Html.Runtime.Worker;

internal sealed class RuntimeViewport(IHtmlDocument document, HtmlScriptRequest options,
    Func<CancellationToken> currentCommandToken, Func<IReadOnlyList<HtmlRuntimeResource>> currentResources) {
    private double _scrollX;
    private double _scrollY;

    internal double ScrollX => _scrollX;
    internal double ScrollY => _scrollY;
    internal double Width => options.ViewportWidth;
    internal double Height => options.ViewportHeight;
    internal bool Enabled => options.Profile == HtmlRuntimeProfile.WebApplicationV1;
    internal CancellationToken CurrentCommandToken => Resolve(CancellationToken.None);

    internal RuntimeElementLayout Measure(IElement element, CancellationToken token) {
        token = Resolve(token);
        if (!Enabled) return RuntimeElementLayout.Unqualified;
        if (RuntimeFocusController.HiddenByMarkup(element))
            return new RuntimeElementLayout(false, false, null, _scrollX, _scrollY, 0D, 0D);
        HtmlInteractionLayoutResult measured = HtmlInteractionLayoutEngine.Measure(
            document, element, Width, Height, options.MaxInputCharacters, options.MaxNodes, options.MaxDepth,
            new Uri(RuntimeDocumentUrls.Base(document)), ResolveStylesheet, token);
        if (!measured.IsConnected || !measured.HasLayoutBox || measured.Bounds is not HtmlInteractionRect bounds)
            return new RuntimeElementLayout(false, measured.AcceptsPointerEvents, null, _scrollX, _scrollY, measured.DocumentWidth, measured.DocumentHeight);
        var box = new HtmlRuntimeRect {
            X = bounds.X - _scrollX,
            Y = bounds.Y - _scrollY,
            Width = bounds.Width,
            Height = bounds.Height
        };
        bool intersects = box.X < Width && box.Y < Height && box.X + box.Width > 0D && box.Y + box.Height > 0D;
        return new RuntimeElementLayout(true, measured.AcceptsPointerEvents, box, _scrollX, _scrollY, measured.DocumentWidth, measured.DocumentHeight, intersects);
    }

    internal RuntimeElementLayout ScrollIntoView(IElement element, CancellationToken token) {
        token = Resolve(token);
        RuntimeElementLayout before = Measure(element, token);
        if (!before.IsVisible || before.Box == null) return before;
        double documentX = before.Box.X + _scrollX;
        double documentY = before.Box.Y + _scrollY;
        if (documentX < _scrollX) _scrollX = documentX;
        else if (documentX + before.Box.Width > _scrollX + Width) _scrollX = documentX + before.Box.Width - Width;
        if (documentY < _scrollY) _scrollY = documentY;
        else if (documentY + before.Box.Height > _scrollY + Height) _scrollY = documentY + before.Box.Height - Height;
        _scrollX = Math.Clamp(_scrollX, 0D, Math.Max(0D, before.DocumentWidth - Width));
        _scrollY = Math.Clamp(_scrollY, 0D, Math.Max(0D, before.DocumentHeight - Height));
        return Measure(element, token);
    }

    internal void ScrollTo(double x, double y, CancellationToken token = default) {
        token = Resolve(token);
        if (!Enabled) return;
        IElement? root = document.Body ?? document.DocumentElement;
        RuntimeElementLayout layout = root == null ? RuntimeElementLayout.Unqualified : Measure(root, token);
        _scrollX = Math.Clamp(double.IsFinite(x) ? x : 0D, 0D, Math.Max(0D, layout.DocumentWidth - Width));
        _scrollY = Math.Clamp(double.IsFinite(y) ? y : 0D, 0D, Math.Max(0D, layout.DocumentHeight - Height));
    }

    internal void ScrollBy(double x, double y, CancellationToken token = default) => ScrollTo(_scrollX + x, _scrollY + y, token);

    private CancellationToken Resolve(CancellationToken token) => token.CanBeCanceled ? token : currentCommandToken();

    private string? ResolveStylesheet(Uri url) {
        if (!url.IsAbsoluteUri || url.Scheme != Uri.UriSchemeHttp && url.Scheme != Uri.UriSchemeHttps || url.UserInfo.Length != 0) return null;
        string key = ResourceKey(url);
        HtmlRuntimeResource? resource = currentResources().LastOrDefault(candidate =>
            ResourceKey(candidate.Url) == key || ResourceKey(candidate.FinalUrl) == key);
        if (resource == null || !resource.ContentType.Split(';', 2)[0].Trim().Equals("text/css", StringComparison.OrdinalIgnoreCase)) return null;
        return HtmlResourcePipeline.TryDecodeStylesheet(resource.Buffer, resource.ContentType, out string css) ? css : null;
    }

    private static string ResourceKey(Uri url) => new UriBuilder(url) { Fragment = string.Empty }.Uri.AbsoluteUri;
}

internal readonly record struct RuntimeElementLayout(
    bool IsVisible,
    bool AcceptsPointerEvents,
    HtmlRuntimeRect? Box,
    double ScrollX,
    double ScrollY,
    double DocumentWidth,
    double DocumentHeight,
    bool IsInViewport = false) {
    internal static RuntimeElementLayout Unqualified => new(false, false, null, 0D, 0D, 0D, 0D);
}
