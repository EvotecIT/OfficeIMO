using System.Globalization;
using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Html.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlNormalizer {
    // HTML serialization cannot encode live control state. Carry it through the existing policy
    // normalizer with private, collision-free markers on an independent tree. Remove every marker
    // before returning the result; source attributes and authored defaults remain unchanged.
    internal static IHtmlDocument NormalizeToDocument(IHtmlDocument source, HtmlNormalizationOptions options) {
        HtmlConversionInputGuard.ValidateDocument(source, options.Limits);
        IElement[] controls = source.QuerySelectorAll("input,textarea,select,option")
            .Where(element => NativeFormState.Get(element) != null).ToArray();
        if (controls.Length == 0) return HtmlDocumentParser.ParseDocument(Normalize(source, options));

        IHtmlDocument prepared = HtmlDocumentParser.CloneDocument(source);
        string marker;
        do { marker = "data-form-state-" + Guid.NewGuid().ToString("N"); }
        while (prepared.QuerySelector("[" + marker + "]") != null);
        var states = new Dictionary<string, (string Name, HtmlFormControlState State)>(StringComparer.Ordinal);
        foreach (IElement element in prepared.QuerySelectorAll("input,textarea,select,option")) {
            if (NativeFormState.Get(element) is not HtmlFormControlState state) continue;
            string key = states.Count.ToString(CultureInfo.InvariantCulture);
            states.Add(key, (element.LocalName, state));
            element.SetAttribute(marker, key);
        }
        IHtmlDocument normalized = HtmlDocumentParser.ParseDocument(Normalize(prepared, options));
        foreach (IElement element in normalized.QuerySelectorAll("[" + marker + "]")) {
            string key = element.GetAttribute(marker)!;
            element.RemoveAttribute(marker);
            if (!states.TryGetValue(key, out var entry) || !string.Equals(element.LocalName, entry.Name, StringComparison.Ordinal))
                throw new InvalidOperationException("Normalization changed a form-state marker's control type.");
            NativeFormState.Attach(element, entry.State);
        }
        NativeFormState.ApplyTree(normalized);
        return normalized;
    }
}
