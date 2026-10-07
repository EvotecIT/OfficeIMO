using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlImageSourceResolver {
    /// <summary>
    /// Resolves the candidates selected by the active render media environment.
    /// </summary>
    internal static HtmlRenderImageSelection SelectImageForRendering(IElement element, Uri? baseUri, HtmlUrlPolicy? policy, HtmlRenderOptions options) {
        var candidates = new CandidateAccumulator();
        if (element == null) return new HtmlRenderImageSelection(candidates.Items, null);

        HtmlResponsiveImageSelectionOptions selectionOptions = CreateResponsiveSelectionOptions(options);

        IElement? selectedPictureSource = null;
        IElement? picture = element.ParentElement;
        if (picture != null && picture.TagName.Equals("PICTURE", StringComparison.OrdinalIgnoreCase)) {
            double mediaWidth = options.CssMediaWidth;
            double mediaHeight = options.CssMediaHeight;
            foreach (IElement child in picture.Children) {
                if (ReferenceEquals(child, element)) break;
                if (!child.TagName.Equals("SOURCE", StringComparison.OrdinalIgnoreCase)
                    || !HtmlComputedStyleEngine.IsApplicableMedia(
                        child.GetAttribute("media") ?? string.Empty,
                        options.MediaContext,
                        mediaWidth,
                        mediaHeight,
                        options.MediaFeatures)
                    || !HtmlPictureSourceSupport.IsSupportedConversionContentType(child.GetAttribute("type"))) {
                    continue;
                }

                int countBeforeSource = candidates.Items.Count;
                AddSelectedResolvedSrcSet(candidates, child, baseUri, policy, selectionOptions, defaultSource: null, SrcSetAttributes);
                if (candidates.Items.Count == countBeforeSource) {
                    int candidateCount = 0;
                    AddResolvedUrlAttributes(candidates, child, baseUri, policy, options.ResponsiveImageCandidateLimit,
                        ref candidateCount, PictureSourceAttributes);
                }
                if (candidates.Items.Count > countBeforeSource) {
                    selectedPictureSource = child;
                    break;
                }
            }
        }

        if (selectedPictureSource == null) {
            AddResolvedUrlAttributes(candidates, element, baseUri, policy, LazySourceAttributes);
            if (candidates.Items.Count == 0) {
                string? defaultSource = element.GetAttribute("src");
                AddSelectedResolvedSrcSet(candidates, element, baseUri, policy, selectionOptions, defaultSource, SrcSetAttributes);
                if (candidates.Items.Count == 0) AddResolvedUrlAttributes(candidates, element, baseUri, policy, SourceAttributes);
            }
        }

        IElement dimensionSource = selectedPictureSource != null
            && (selectedPictureSource.HasAttribute("width") || selectedPictureSource.HasAttribute("height"))
            ? selectedPictureSource : element;
        return new HtmlRenderImageSelection(candidates.Items, dimensionSource);
    }

}

/// <summary>Pairs render-media image candidates with their HTML dimension-hint owner.</summary>
internal readonly struct HtmlRenderImageSelection {
    internal HtmlRenderImageSelection(IReadOnlyList<string> sources, IElement? dimensionSource) {
        Sources = sources;
        DimensionSource = dimensionSource;
    }

    internal IReadOnlyList<string> Sources { get; }
    internal IElement? DimensionSource { get; }
}
