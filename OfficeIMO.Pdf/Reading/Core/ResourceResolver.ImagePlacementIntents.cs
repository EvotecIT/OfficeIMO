using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class ResourceResolver {
    private static List<PdfImagePlacement> GetDistinctImageMaskPlacements(
        IReadOnlyList<PdfImagePlacement> placements,
        bool hasAuthoredImageIntent,
        OfficeIccRenderingIntent authoredImageIntent) {
        var effectivePlacements = new List<PdfImagePlacement>();
        var indexes = new Dictionary<string, int>(System.StringComparer.Ordinal);
        for (int index = 0; index < placements.Count; index++) {
            PdfImagePlacement placement = placements[index];
            OfficeIccRenderingIntent renderingIntent = hasAuthoredImageIntent
                ? authoredImageIntent
                : placement.RenderingIntent;
            string key = BuildImageMaskIntentKey(placement.ImageMaskColor, renderingIntent);
            if (indexes.TryGetValue(key, out int existingIndex)) {
                if (!effectivePlacements[existingIndex].HasAuthoredRenderingIntent &&
                    placement.HasAuthoredRenderingIntent) {
                    effectivePlacements[existingIndex] = placement;
                }
                continue;
            }
            indexes[key] = effectivePlacements.Count;
            effectivePlacements.Add(placement);
        }
        return effectivePlacements;
    }

    private static List<EffectiveImageIntent> GetDistinctImageIntents(
        IReadOnlyList<PdfImagePlacement> placements,
        bool hasAuthoredImageIntent,
        OfficeIccRenderingIntent authoredImageIntent) {
        var effectiveIntents = new List<EffectiveImageIntent>();
        var indexes = new Dictionary<OfficeIccRenderingIntent, int>();
        for (int index = 0; index < placements.Count; index++) {
            PdfImagePlacement placement = placements[index];
            OfficeIccRenderingIntent renderingIntent = hasAuthoredImageIntent
                ? authoredImageIntent
                : placement.RenderingIntent;
            bool authored = hasAuthoredImageIntent || placement.HasAuthoredRenderingIntent;
            if (indexes.TryGetValue(renderingIntent, out int existingIndex)) {
                if (authored && !effectiveIntents[existingIndex].HasAuthoredRenderingIntent) {
                    effectiveIntents[existingIndex] = new EffectiveImageIntent(renderingIntent, true);
                }
                continue;
            }
            indexes[renderingIntent] = effectiveIntents.Count;
            effectiveIntents.Add(new EffectiveImageIntent(renderingIntent, authored));
        }
        return effectiveIntents;
    }

    private static string BuildImageMaskIntentKey(
        OfficeColor imageMaskColor,
        OfficeIccRenderingIntent renderingIntent) =>
        imageMaskColor.R.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," +
        imageMaskColor.G.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," +
        imageMaskColor.B.ToString(System.Globalization.CultureInfo.InvariantCulture) + "," +
        imageMaskColor.A.ToString(System.Globalization.CultureInfo.InvariantCulture) +
        "|intent:" + ((int)renderingIntent).ToString(System.Globalization.CultureInfo.InvariantCulture);

    private static string BuildImageResourceKey(
        int pageNumber,
        string resourceName,
        int objectNumber,
        int directStreamIdentity,
        OfficeColor imageMaskColor,
        OfficeIccRenderingIntent renderingIntent) =>
        BuildImageResourceKey(pageNumber, resourceName, objectNumber, directStreamIdentity, imageMaskColor) +
        "|intent:" +
        ((int)renderingIntent).ToString(System.Globalization.CultureInfo.InvariantCulture);

    private readonly struct EffectiveImageIntent {
        internal EffectiveImageIntent(
            OfficeIccRenderingIntent renderingIntent,
            bool hasAuthoredRenderingIntent) {
            RenderingIntent = renderingIntent;
            HasAuthoredRenderingIntent = hasAuthoredRenderingIntent;
        }

        internal OfficeIccRenderingIntent RenderingIntent { get; }
        internal bool HasAuthoredRenderingIntent { get; }
    }
}
