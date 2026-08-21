using AngleSharp.Dom;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Html;

public static partial class HtmlEditableLayoutProjector {
    private static readonly ConditionalWeakTable<IElement, EditableLayoutMarker> EditableLayoutMarkers = new();

    internal static void SetRegionSourceKey(IElement element, string sourceKey) =>
        EditableLayoutMarkers.GetOrCreateValue(element).RegionSourceKey = sourceKey;

    internal static string? GetRegionSourceKey(IElement element) =>
        EditableLayoutMarkers.TryGetValue(element, out EditableLayoutMarker? marker)
            ? marker.RegionSourceKey
            : null;

    internal static void SetImageSourceKey(IElement element, string sourceKey) =>
        EditableLayoutMarkers.GetOrCreateValue(element).ImageSourceKey = sourceKey;

    internal static string? GetImageSourceKey(IElement element) =>
        EditableLayoutMarkers.TryGetValue(element, out EditableLayoutMarker? marker)
            ? marker.ImageSourceKey
            : null;

    internal static void CopyMarkers(IDocument source, IDocument target) {
        IReadOnlyList<IElement> sourceElements = source.QuerySelectorAll("*").ToList();
        IReadOnlyList<IElement> targetElements = target.QuerySelectorAll("*").ToList();
        int count = Math.Min(sourceElements.Count, targetElements.Count);
        for (int index = 0; index < count; index++) {
            if (!EditableLayoutMarkers.TryGetValue(sourceElements[index], out EditableLayoutMarker? marker)) continue;
            EditableLayoutMarker targetMarker = EditableLayoutMarkers.GetOrCreateValue(targetElements[index]);
            targetMarker.RegionSourceKey = marker.RegionSourceKey;
            targetMarker.ImageSourceKey = marker.ImageSourceKey;
        }
    }

    private static void RestoreAuthoredAttribute(
        IElement element,
        string attributeName,
        string? sourceKey,
        IReadOnlyDictionary<string, string?> authoredValues) {
        if (sourceKey == null || !authoredValues.TryGetValue(sourceKey, out string? authoredValue)) return;
        if (authoredValue == null) element.RemoveAttribute(attributeName);
        else element.SetAttribute(attributeName, authoredValue);
    }

    private sealed class EditableLayoutMarker {
        internal string? RegionSourceKey { get; set; }
        internal string? ImageSourceKey { get; set; }
    }
}