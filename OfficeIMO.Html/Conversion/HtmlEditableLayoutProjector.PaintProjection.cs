namespace OfficeIMO.Html;

public static partial class HtmlEditableLayoutProjector {
    private static List<EditableLayoutRegionOccurrence> CoalescePaintProjectionOccurrences(
        IReadOnlyList<EditableLayoutRegionOccurrence> occurrences) {
        var result = new List<EditableLayoutRegionOccurrence>();
        foreach (var group in occurrences.GroupBy(occurrence => new {
            // Unprojected occurrences remain distinct even if a scene references
            // the same visual twice. Only explicitly split paint shares identity.
            Identity = occurrence.Region.PaintProjectionIdentity ?? new object(),
            occurrence.Region.SourceKey,
            occurrence.Page,
            occurrence.Region.X,
            occurrence.Region.Y,
            occurrence.Region.Width,
            occurrence.Region.Height,
            occurrence.Region.LayoutY,
            occurrence.Region.LayoutHeight,
            occurrence.SectionOriginX,
            occurrence.SectionOriginY,
            occurrence.TableOriginX,
            occurrence.TableOriginY
        })) {
            EditableLayoutRegionOccurrence first = group.First();
            if (group.Count() == 1) {
                result.Add(first);
                continue;
            }
            // Page, normal-flow bounds and semantic origins must all agree.
            // Genuine page/column fragments stay separate and are diagnosed by
            // the existing occurrence contract; paint splitting adds no region.
            HtmlRenderLayoutRegion merged = (HtmlRenderLayoutRegion)first.Region.ProjectPaint(
                group.OrderBy(item => item.Region.PaintOrder).SelectMany(item => item.Region.Visuals),
                0D, 0D, group.Min(item => item.Region.PaintOrder));
            merged.IdentifyPaintProjection(group.Key.Identity);
            result.Add(new EditableLayoutRegionOccurrence(first.Page, merged,
                first.SectionOriginX, first.SectionOriginY, first.TableOriginX, first.TableOriginY));
        }
        return result;
    }
}
