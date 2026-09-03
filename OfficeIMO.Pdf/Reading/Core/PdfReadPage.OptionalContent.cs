namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    private static readonly System.Threading.AsyncLocal<Action?> HiddenOptionalContentInspectionObserver =
        new System.Threading.AsyncLocal<Action?>();

    internal static Action? HiddenOptionalContentInspectionObserverForTesting {
        get => HiddenOptionalContentInspectionObserver.Value;
        set => HiddenOptionalContentInspectionObserver.Value = value;
    }

    internal bool IsHiddenOptionalContent(PdfDictionary? sourceDictionary) {
        if (sourceDictionary is null ||
            !sourceDictionary.Items.TryGetValue("OC", out PdfObject? optionalContentObject)) {
            return false;
        }

        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        PdfPageOptionalContentVisibility? visibility = GetOptionalContentVisibility(pageResources);
        return visibility?.HasUnsupportedViewUsageApplications == false &&
            visibility.IsHidden(optionalContentObject);
    }

    internal bool HasUnsupportedOptionalContentViewUsageApplications() {
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        return GetOptionalContentVisibility(pageResources)?.HasUnsupportedViewUsageApplications == true;
    }

    internal IReadOnlyList<PdfTextSpan> GetHiddenOptionalContentTextSpans(bool includeArtifactText) {
        HiddenOptionalContentInspectionObserver.Value?.Invoke();
        if (HasUnsupportedOptionalContentViewUsageApplications()) {
            return Array.Empty<PdfTextSpan>();
        }

        IReadOnlyList<PdfTextSpan> visible = GetTextSpans(includeArtifactText, default);
        IReadOnlyList<PdfTextSpan> includingHidden = GetTextSpans(
            includeArtifactText,
            default,
            includeHiddenOptionalContent: true);
        var visibleKeys = new HashSet<PdfContentOrderKey>(visible
            .Select(static span => span.ContentOrderKey)
            .OfType<PdfContentOrderKey>());
        return includingHidden
            .Where(span => span.ContentOrderKey is not null && !visibleKeys.Contains(span.ContentOrderKey))
            .ToArray();
    }
}
