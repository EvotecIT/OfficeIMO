using OfficeIMO.Drawing;
using static OfficeIMO.Publisher.Internal.PublisherTextComparison;

namespace OfficeIMO.Publisher.Internal;

internal sealed partial class PublisherNativeWriter {
    private void Compare() {
        if (!MapEqual(_baseline.Text.Stories, _current.Text.Stories, StoryEqual)
            || !MapEqual(_baseline.Text.TableStories, _current.Text.TableStories, TableStoryEqual)) {
            var sourceStyles = _imported.Styles.Stories.ToDictionary(story => story.StoryId);
            foreach (var story in _baseline.Text.Stories) {
                _token.ThrowIfCancellationRequested();
                if ((!_current.Text.Stories.TryGetValue(story.Key, out PublisherTextStory? current) || !StoryEqual(story.Value, current)
                    || !TableStoryUnchanged(story.Key))
                    && sourceStyles.TryGetValue(story.Key, out PublisherQuillStoryStyles? dependency)) _styleDependencies.Add(dependency);
            }
            if (!TryRewriteStories()) Reject("Quill", "Stories or formatting changed outside the supported existing-story and formatting-page profiles. StyleDependencies identifies the original stories, definition ordinals and direct native OPL spans.");
        }
        if (!PagesEqual() || !MapEqual(_baseline.ObjectOwners, _current.ObjectOwners, (a, b) => a == b)
            || !SameKeys(_baseline.Text.Frames, _current.Text.Frames) || !SameKeys(_baseline.Shapes, _current.Shapes)
            || !SameKeys(_baseline.Pictures, _current.Pictures) || !SameKeys(_baseline.Tables, _current.Tables)) {
            Reject("Contents", "Page/master topology, object ownership, order or object identities changed. Authored objects have no native records in this operation."); return;
        }
        if (!MapEqual(_baseline.Text.Chains, _current.Text.Chains, (a, b) => a.SequenceEqual(b))
            || !MapEqual(_baseline.Text.TableStories, _current.Text.TableStories, (a, b) => ReferenceEquals(a, b)
                || a.Segments.Count == b.Segments.Count && (_quillBytes != null || a.Segments.Zip(b.Segments, StoryEqual).All(value => value)))
            || !MapEqual(_baseline.TableStoryReferences, _current.TableStoryReferences, (a, b) => a == b))
            Reject("Quill", "Frame chains or table story ownership changed. Quill rewriting is unavailable.");
        if (!MapEqual(_baseline.Text.WrapInfo, _current.Text.WrapInfo, (a, b) => Equals(a, b))
            || !MapEqual(_baseline.Text.ImportedObstacles, _current.Text.ImportedObstacles, (a, b) => Equals(a, b)))
            Reject("Contents", "Native wrap intent or unsupported artwork placement changed.");
        PrepareImages();
        foreach (var entry in _baseline.Tables) {
            _token.ThrowIfCancellationRequested();
            if (!TableEqual(entry.Value, _current.Tables[entry.Key]))
                Reject("Contents/object/" + entry.Key, "Table placement, tracks, cells or paint changed. Native table rewriting is unavailable.");
        }
        foreach (var entry in _baseline.Text.Frames) CompareFrame(entry.Value, _current.Text.Frames[entry.Key]);
        foreach (var entry in _baseline.Pictures) ComparePicture(entry.Value, _current.Pictures[entry.Key]);
        foreach (var entry in _baseline.Shapes) CompareShape(entry.Value, _current.Shapes[entry.Key]);
    }

    private bool TryRewriteStories() {
        if (!SameKeys(_baseline.Text.Stories, _current.Text.Stories)
            || !SameKeys(_baseline.Text.TableStories, _current.Text.TableStories)) return false;
        bool extended = false;
        foreach (var entry in _baseline.Text.Stories) {
            PublisherTextStory current = _current.Text.Stories[entry.Key];
            if (StoryEqual(entry.Value, current) && TableStoryUnchanged(entry.Key)) continue;
            bool changedFormatting;
            if (_baseline.Text.TableStories.ContainsKey(entry.Key)) {
                if (!CanRewriteTableStory(entry.Key)) return false;
                changedFormatting = true;
            } else if (!CanRewriteStory(entry.Value, current, out changedFormatting) || !CanRewriteStoryFrames(entry.Key)) return false;
            extended |= changedFormatting;
        }
        if (!extended && _baseline.Text.Stories.Count == 1 && _baseline.Text.Frames.Count == 1 && _styleDependencies.Count == 1) {
            _quillBytes = PublisherQuillTextWriter.TryRewrite(_imported.NativeSource!.QuillBytes, _styleDependencies[0],
                _baseline.Text.Stories.Values.Single(), _current.Text.Stories.Values.Single(), _sourceLimits, _maximumBytes, _token);
            if (_quillBytes != null) return true;
        }
        _quillBytes = PublisherQuillStoriesWriter.TryRewrite(_imported.NativeSource!.QuillBytes, _imported.Styles.Stories,
            _baseline.Text.Stories, _current.Text.Stories, _baseline.Text.TableStories, _current.Text.TableStories,
            _imported.Contents.Palette, _sourceLimits, _maximumBytes, _token);
        return _quillBytes != null;
    }

    private bool CanRewriteTableStory(uint storyId) {
        PublisherTableState[] tables = _baseline.Tables.Values.Where(table => table.StoryId == storyId).ToArray();
        if (tables.Length == 0 || !tables.All(table => _baseline.CellTextEditLimitation(table) == null
            && _current.Tables.TryGetValue(table.Id, out PublisherTableState? current) && current.StoryId == storyId
            && _current.CellTextEditLimitation(current) == null)) return false;
        PublisherTableStory before = _baseline.Text.TableStories[storyId], after = _current.Text.TableStories[storyId];
        return before.Segments.Count == after.Segments.Count && before.Segments.Zip(after.Segments, (a, b) =>
            StoryEqual(a, b) || CanRewriteStory(a, b, out _)).All(value => value);
    }

    private bool TableStoryUnchanged(uint storyId) => _baseline.Text.TableStories.TryGetValue(storyId, out PublisherTableStory? before)
        ? _current.Text.TableStories.TryGetValue(storyId, out PublisherTableStory? after) && TableStoryEqual(before, after)
        : !_current.Text.TableStories.ContainsKey(storyId);

    private static bool TableStoryEqual(PublisherTableStory a, PublisherTableStory b) => ReferenceEquals(a, b)
        || a.Segments.Count == b.Segments.Count && a.Segments.Zip(b.Segments, StoryEqual).All(value => value);

    private bool CanRewriteStoryFrames(uint storyId) {
        int frameCount = _baseline.Text.Frames.Values.Count(frame => frame.StoryId == storyId);
        if (frameCount == 1) return true;
        if (frameCount == 0 || !_baseline.Text.Chains.TryGetValue(storyId, out IReadOnlyList<uint>? chain)
            || chain.Count != frameCount) return false;
        var pages = new HashSet<uint>(_baseline.Pages.Where(page => !page.IsMaster).Select(page => page.Id));
        return chain.All(id => _baseline.Text.Frames.TryGetValue(id, out PublisherFrameState? frame)
            && frame.StoryId == storyId && frame.Available && frame.SupportsContinuation
            && frame.PageId.HasValue && pages.Contains(frame.PageId.Value));
    }

    private static bool CanRewriteStory(PublisherTextStory before, PublisherTextStory after, out bool extended) {
        extended = false;
        if (before.Id != after.Id || before.Paragraphs.Count == 0 || after.Paragraphs.Count == 0) return false;
        bool reshape = before.Paragraphs.Count != after.Paragraphs.Count;
        if (reshape && before.Paragraphs.Any(paragraph => paragraph.Runs.Count != 1 || paragraph.Label != null
            || !ParagraphFormattingEqual(before.Paragraphs[0], paragraph)
            || !RunFormattingEqual(before.Paragraphs[0].Runs[0], paragraph.Runs[0]))) {
            // Formatting matches admit text replacement; the native planner
            // separately requires complete, unambiguous source definitions.
            var retained = new HashSet<OfficeRichTextParagraph>(before.Paragraphs.Where(p => p.Label == null), PublisherTextComparison.Paragraphs);
            var formats = new HashSet<OfficeRichTextParagraph>(before.Paragraphs.Where(p => p.Label == null
                && p.Runs.Count == 1 && p.Runs[0].Text.Length > 0), PublisherTextComparison.ParagraphFormats);
            extended = true;
            return after.Paragraphs.All(p => p.Label == null && (retained.Contains(p)
                || p.Runs.Count == 1 && p.Runs[0].Text.Length > 0 && formats.Contains(p)));
        }
        for (int paragraph = 0; paragraph < after.Paragraphs.Count; paragraph++) {
            OfficeRichTextParagraph a = before.Paragraphs[reshape ? 0 : paragraph], b = after.Paragraphs[paragraph];
            bool empty = b.Runs.Count == 0 || b.Runs.Count == 1 && b.Runs[0].Text.Length == 0;
            if (a.Runs.Count < 1 || !empty && a.Runs.Count != 1 && a.Runs.Count != b.Runs.Count
                || a.Label != null || b.Label != null) return false;
            extended |= empty || PublisherQuillParagraphStyleWriter.HasChanges(a, b);
            if (empty) {
                if (b.Runs.Count == 1 && !RunPreservedFormattingEqual(a.Runs[a.Runs.Count - 1], b.Runs[0])) return false;
                continue;
            }
            for (int index = 0; index < b.Runs.Count; index++) {
                if (b.Runs[index].Text.Length == 0
                    || !RunPreservedFormattingEqual(a.Runs[a.Runs.Count == 1 ? 0 : index], b.Runs[index])) return false;
                extended |= PublisherQuillCharacterStyleWriter.HasExtendedChanges(a.Runs[a.Runs.Count == 1 ? 0 : index], b.Runs[index]);
            }
        }
        return true;
    }

    private bool PagesEqual() => _baseline.Pages.Count == _current.Pages.Count
        && _baseline.Pages.Zip(_current.Pages, (a, b) => ReferenceEquals(a, b)
            || a.Id == b.Id && a.Name == b.Name && a.IsMaster == b.IsMaster && a.HasNativeOrigin == b.HasNativeOrigin
            && a.Width == b.Width && a.Height == b.Height && a.MasterPageId == b.MasterPageId
            && a.RegistrationX == b.RegistrationX && a.RegistrationY == b.RegistrationY
            && a.SupportsPlacementEditing == b.SupportsPlacementEditing && a.HasUnassessedBackground == b.HasUnassessedBackground
            && a.FrameEditRange == b.FrameEditRange && a.ObjectIds.SequenceEqual(b.ObjectIds)).All(value => value);

    private void CompareFrame(PublisherFrameState a, PublisherFrameState b) {
        _token.ThrowIfCancellationRequested();
        if (ReferenceEquals(a, b)) return;
        if (a.Id != b.Id || a.StoryId != b.StoryId || a.PageId != b.PageId || a.HasAnchor != b.HasAnchor
            || a.TransformAvailable != b.TransformAvailable || a.Hidden != b.Hidden || a.Drawable != b.Drawable
            || a.CanEditBounds != b.CanEditBounds || a.SupportsContinuation != b.SupportsContinuation || a.IsAuthored != b.IsAuthored
            || a.LayoutEditLimitation != b.LayoutEditLimitation || a.EditRange != b.EditRange || !a.WrapObjectIds.SequenceEqual(b.WrapObjectIds)) {
            Reject("Contents/object/" + a.Id, "A frame input other than its bounds or layout changed."); return;
        }
        FrameLayout(a, b);
        if (SameBounds(a.X, a.Y, a.Width, a.Height, b.X, b.Y, b.Width, b.Height) && a.PageTransform == b.PageTransform) return;
        if (!a.CanEditBounds || b.PageTransform != OfficeTransform.Translate(b.X, b.Y)) {
            Reject("Contents/object/" + a.Id, "The frame has no qualified untransformed native bounds edit."); return;
        }
        Bounds(a.Id, a.PageId, b.X, b.Y, b.Width, b.Height);
    }

    private void ComparePicture(PublisherPicture a, PublisherPicture b) {
        _token.ThrowIfCancellationRequested();
        if (ReferenceEquals(a, b)) return;
        if (a.Id != b.Id || a.PageId != b.PageId || a.ImageId != b.ImageId && !_pictureNames.ContainsKey(a.Id) || !a.Crop.Equals(b.Crop)
            || a.HasPlacement != b.HasPlacement || a.CanEdit != b.CanEdit || a.EditLimitation != b.EditLimitation || a.IsAuthored != b.IsAuthored
            || !a.CanEdit || !ShapeEqual(PublisherShapeGeometry.Resize(a.CreatePaint(), b.Width, b.Height), b.CreatePaint())) {
            Reject("Contents/object/" + a.Id, "Picture asset binding, crop, paint or unsupported placement changed."); return;
        }
        if (SameBounds(a.X, a.Y, a.Width, a.Height, b.X, b.Y, b.Width, b.Height) && a.PageTransform == b.PageTransform) return;
        if (b.PageTransform != OfficeTransform.Translate(b.X, b.Y)) {
            Reject("Contents/object/" + a.Id, "The picture has no qualified untransformed native bounds edit."); return;
        }
        Bounds(a.Id, a.PageId, b.X, b.Y, b.Width, b.Height);
    }

    private void CompareShape(PublisherShape a, PublisherShape b) {
        _token.ThrowIfCancellationRequested();
        if (ReferenceEquals(a, b)) return;
        if (a.Id != b.Id || a.PageId != b.PageId || a.HasPlacement != b.HasPlacement || a.EditLimitation != b.EditLimitation
            || a.IsAuthored != b.IsAuthored || !a.CanEdit || !b.CanEdit) {
            Reject("Contents/object/" + a.Id, "An unsupported vector-artwork input changed."); return;
        }
        // A text-capable shape's canonical geometry belongs to its frame.
        PublisherShape before = _baseline.CurrentShape(a.Id), after = _current.CurrentShape(b.Id);
        OfficeShape prior = before.CreateShape(), current = after.CreateShape();
        if (!GeometryEqual(PublisherShapeGeometry.Resize(prior, after.Width, after.Height), current)) {
            Reject("Contents/object/" + a.Id, "Vector geometry or effects changed beyond the fixed solid-paint slice."); return;
        }
        if (!SameBounds(before.X, before.Y, before.Width, before.Height, after.X, after.Y, after.Width, after.Height)
            || before.PageTransform != after.PageTransform) {
            if (after.PageTransform != OfficeTransform.Translate(after.X, after.Y)) {
                Reject("Contents/object/" + a.Id, "The vector artwork has no qualified untransformed native bounds edit."); return;
            }
            Bounds(a.Id, a.PageId, after.X, after.Y, after.Width, after.Height);
        }
        if (!PaintEqual(prior, current)) Paint(a.Id, prior, current);
    }

    private static bool TableEqual(PublisherTableState a, PublisherTableState b) => ReferenceEquals(a, b)
        || a.Id == b.Id && a.StoryId == b.StoryId && a.PageId == b.PageId && a.X == b.X && a.Y == b.Y
        && a.Width == b.Width && a.Height == b.Height && a.PageTransform == b.PageTransform && a.HasPlacement == b.HasPlacement
        && a.Hidden == b.Hidden && a.Padding.Equals(b.Padding) && a.BoundsEditLimitation == b.BoundsEditLimitation && a.IsAuthored == b.IsAuthored
        && a.Columns.SequenceEqual(b.Columns) && a.Rows.SequenceEqual(b.Rows) && a.Cells.SequenceEqual(b.Cells)
        && ShapeEqual(a.CreatePaint(), b.CreatePaint());

    private static bool SameBounds(double ax, double ay, double aw, double ah, double bx, double by, double bw, double bh) =>
        ax == bx && ay == by && aw == bw && ah == bh;
    private static bool SameKeys<TKey, TValue>(IReadOnlyDictionary<TKey, TValue> a, IReadOnlyDictionary<TKey, TValue> b) where TKey : notnull =>
        a.Count == b.Count && a.Keys.All(b.ContainsKey);
    private static bool MapEqual<TKey, TValue>(IReadOnlyDictionary<TKey, TValue> a, IReadOnlyDictionary<TKey, TValue> b,
        Func<TValue, TValue, bool> equal) where TKey : notnull => ReferenceEquals(a, b)
        || SameKeys(a, b) && a.All(item => equal(item.Value, b[item.Key]));

    private static bool PaintEqual(OfficeShape a, OfficeShape b) => a.FillColor == b.FillColor && a.StrokeColor == b.StrokeColor
        && a.StrokeWidth == b.StrokeWidth && (a.FillOpacity ?? 1) == (b.FillOpacity ?? 1) && (a.StrokeOpacity ?? 1) == (b.StrokeOpacity ?? 1);
    private static bool ShapeEqual(OfficeShape? a, OfficeShape? b) => a == null || b == null ? a == b : GeometryEqual(a, b) && PaintEqual(a, b);
    private static bool GeometryEqual(OfficeShape a, OfficeShape b) => a.Kind == b.Kind && a.Width == b.Width && a.Height == b.Height
        && a.CornerRadius == b.CornerRadius && a.Points.SequenceEqual(b.Points) && a.PathCommands.SequenceEqual(b.PathCommands)
        && a.StrokeDashStyle == b.StrokeDashStyle && a.StrokeDashOffset == b.StrokeDashOffset && a.StrokeDashArray.SequenceEqual(b.StrokeDashArray)
        && a.StrokeLineCap == b.StrokeLineCap && a.StrokeLineJoin == b.StrokeLineJoin && a.StrokeMiterLimit == b.StrokeMiterLimit
        && a.StrokeHairline == b.StrokeHairline && a.CrispEdges == b.CrispEdges && a.FillRule == b.FillRule && a.Transform == b.Transform
        && a.FillGradient == null && b.FillGradient == null && a.FillRadialGradient == null && b.FillRadialGradient == null
        && a.StrokeGradient == null && b.StrokeGradient == null && a.StrokeRadialGradient == null && b.StrokeRadialGradient == null
        && a.Shadow == null && b.Shadow == null && a.Glow == null && b.Glow == null && a.ClipPath == null && b.ClipPath == null
        && a.StrokeStartMarker == null && b.StrokeStartMarker == null && a.StrokeEndMarker == null && b.StrokeEndMarker == null;

}
