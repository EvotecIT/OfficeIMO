using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        private readonly List<(PowerPointShape Shape, PowerPointLayoutBox Applied)> _authoringContent = new();
        private PowerPointLayoutSlide? _authoringHeader;

        private PowerPointLayoutContext AuthoringContext() {
            PowerPointDocument document = OwnerDocument ?? throw new InvalidOperationException("Automatic placement requires an attached slide.");
            return new PowerPointLayoutContext(document, document.LayoutTheme);
        }

        private PowerPointLayoutSlide AuthoringHeader(string? title = null) {
            if (_authoringHeader != null && title == null) return _authoringHeader;
            var header = new PowerPointLayoutSlide { Title = title, Layout = PowerPointLayout.TitleAndContent };
            PowerPointSlideLayoutInfo? layout = GetLayoutInfo();
            if (layout.HasValue) header.TemplateLayout = OwnerDocument?.InspectTemplate().Masters
                .SelectMany(master => master.Layouts)
                .FirstOrDefault(item => item.MasterIndex == layout.Value.MasterIndex && item.LayoutIndex == layout.Value.LayoutIndex);
            return header;
        }

        private double AuthoringTitleBottom() => TextBoxes
            .Where(shape => !shape.Hidden && (shape.PlaceholderType == PowerPointPlaceholderType.Title || shape.PlaceholderType == PowerPointPlaceholderType.CenteredTitle))
            .Select(shape => shape.TryGetExportBoundsPoints(out _, out double top, out _, out double height) ? top + height : 0D)
            .DefaultIfEmpty(0D).Max();

        private PowerPointLayoutBox GetAuthoringTitleBounds(string text) {
            PowerPointLayoutContext context = AuthoringContext();
            PowerPointLayoutSlide content = AuthoringHeader(text);
            PowerPointLayoutEngine.HeaderGeometry header = PowerPointLayoutEngine.MeasureHeader(context, content);
            PowerPointLayoutBox? template = PowerPointLayoutEngine.TemplateTitleBounds(context, content);
            return template ?? PowerPointLayoutEngine.PhysicalBox(new PowerPointLayoutRect(context.Grid.X(1), header.TitleY,
                context.Grid.W(10), header.TitleLines * header.TitleSize * context.Style(PowerPointTextRole.Title).LineHeight + 8D));
        }

        private PowerPointTextBox AddAuthoringTitle(string text) {
            if (TextBoxes.Any(shape => shape.PlaceholderType == PowerPointPlaceholderType.Title || shape.PlaceholderType == PowerPointPlaceholderType.CenteredTitle))
                throw new InvalidOperationException("Automatic placement supports one title. Use explicit geometry for additional titles.");
            PowerPointLayoutContext context = AuthoringContext();
            PowerPointLayoutSlide header = AuthoringHeader(text);
            // Validate the measured title/content combination before mutating the native shape tree.
            PowerPointLayoutEngine.AuthoringFrames(context, header, Math.Max(1, _authoringContent.Count), 0D);
            int originalCount = Shapes.Count;
            try {
                PowerPointLayoutEngine.RenderHeader(context, new PowerPointLayoutWriter(this, context.Theme), header, decorate: false);
                PowerPointTextBox title = TextBoxes.LastOrDefault(shape => shape.PlaceholderType == PowerPointPlaceholderType.Title ||
                    shape.PlaceholderType == PowerPointPlaceholderType.CenteredTitle)
                    ?? new PowerPointLayoutWriter(this, context.Theme).Text("Title",
                        new PowerPointLayoutRect(context.Grid.X(1), context.Grid.Y(2), context.Grid.W(10),
                            context.Style(PowerPointTextRole.Title).Size * context.Style(PowerPointTextRole.Title).LineHeight + 8D),
                        context.Paragraph(text, PowerPointTextRole.Title), placeholder: "<p:ph type=\"title\"/>");
                _authoringHeader = header;
                ReflowAuthoringContent();
                return title;
            } catch {
                foreach (PowerPointShape shape in Shapes.Skip(originalCount).ToArray()) RemoveShape(shape);
                _authoringHeader = null;
                throw;
            }
        }

        private void PreflightTitleReflow(OfficeLength top, OfficeLength height) {
            PruneAuthoringContent();
            if (_authoringContent.Count == 0) return;
            PowerPointLayoutEngine.AuthoringFrames(AuthoringContext(), AuthoringHeader(), _authoringContent.Count,
                Math.Max(AuthoringTitleBottom(), (top + height).ToPoints()));
        }

        private PowerPointLayoutBox GetAuthoringContentBounds(bool automatic) {
            if (automatic) {
                PruneAuthoringContent();
                if (_authoringContent.Count >= 2)
                    throw new InvalidOperationException("Automatic placement supports two content objects. Use another slide or explicit geometry.");
            }
            PowerPointLayoutRect[] frames = PowerPointLayoutEngine.AuthoringFrames(AuthoringContext(), AuthoringHeader(),
                automatic ? _authoringContent.Count + 1 : 1, AuthoringTitleBottom());
            return PowerPointLayoutEngine.PhysicalBox(frames[automatic ? _authoringContent.Count : 0]);
        }

        private void AttachAuthoringContent(PowerPointShape shape) {
            _authoringContent.Add((shape, shape.Bounds));
            ReflowAuthoringContent();
        }

        /// <summary>Caller geometry edits and removals end automatic ownership; the native objects remain authoritative.</summary>
        private void PruneAuthoringContent() => _authoringContent.RemoveAll(entry =>
            !Shapes.Contains(entry.Shape) || !SameAuthoringBounds(entry.Shape.Bounds, entry.Applied));

        private static bool SameAuthoringBounds(PowerPointLayoutBox left, PowerPointLayoutBox right) =>
            left.Left == right.Left && left.Top == right.Top && left.Width == right.Width && left.Height == right.Height;

        private void ReflowAuthoringContent() {
            PruneAuthoringContent();
            if (_authoringContent.Count == 0) return;
            PowerPointLayoutRect[] frames = PowerPointLayoutEngine.AuthoringFrames(AuthoringContext(), AuthoringHeader(),
                _authoringContent.Count, AuthoringTitleBottom());
            for (int i = 0; i < _authoringContent.Count; i++) {
                PowerPointShape shape = _authoringContent[i].Shape;
                double[]? columns = shape is PowerPointTable table
                    ? Enumerable.Range(0, table.Columns).Select(column => (double)table.GetColumnWidth(column).ToEmus()).ToArray() : null;
                double[]? rows = shape is PowerPointTable rowTable
                    ? Enumerable.Range(0, rowTable.Rows).Select(row => (double)rowTable.GetRowHeight(row).ToEmus()).ToArray() : null;
                shape.Bounds = PowerPointLayoutEngine.PhysicalBox(frames[i]);
                if (shape is PowerPointTable resized) {
                    resized.SetColumnWidthsByRatio(columns!);
                    resized.SetRowHeightsByRatio(rows!);
                }
                _authoringContent[i] = (shape, shape.Bounds);
            }
        }
    }
}
