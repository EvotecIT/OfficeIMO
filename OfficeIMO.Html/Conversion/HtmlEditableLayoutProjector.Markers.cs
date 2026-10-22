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
            CopyMarkers(sourceElements[index], targetElements[index]);
        }
    }

    internal static void CopyMarkers(IElement source, IElement target) {
        if (!EditableLayoutMarkers.TryGetValue(source, out EditableLayoutMarker? marker)) return;
        EditableLayoutMarker targetMarker = EditableLayoutMarkers.GetOrCreateValue(target);
        targetMarker.RegionSourceKey = marker.RegionSourceKey;
        targetMarker.ImageSourceKey = marker.ImageSourceKey;
    }

    private sealed class EditableLayoutMarker {
        internal string? RegionSourceKey { get; set; }
        internal string? ImageSourceKey { get; set; }
    }
}
