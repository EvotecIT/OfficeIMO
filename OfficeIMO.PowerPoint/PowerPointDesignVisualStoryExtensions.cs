using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public static partial class PowerPointDesignExtensions {
        /// <summary>Adds a screenshot story with semantic crop, alternative text, caption, provenance, and annotations.</summary>
        public static PowerPointSlide AddDesignerScreenshotStorySlide(this PowerPointPresentation presentation,
            string title, string? subtitle, PowerPointImageAsset image, IEnumerable<string>? narrative = null,
            PowerPointDesignTheme? theme = null, PowerPointScreenshotStorySlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (image == null) throw new ArgumentNullException(nameof(image));
            image.Validate();
            List<string> points = (narrative ?? Array.Empty<string>())
                .Where(value => !string.IsNullOrWhiteSpace(value)).ToList();
            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointScreenshotStorySlideOptions resolved = options ?? new PowerPointScreenshotStorySlideOptions();
            PowerPointSlide slide = AddDesignerSlide(presentation, resolved);
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            PrepareLightStorySlide(slide, resolvedTheme, resolved, title, subtitle, width, height);
            PowerPointScreenshotStoryLayoutVariant variant = ResolveScreenshotVariant(resolved, image, points);
            if (variant == PowerPointScreenshotStoryLayoutVariant.SplitNarrative) {
                AddScreenshotSplitNarrative(slide, resolvedTheme, image, points, width, height);
            } else {
                AddScreenshotHero(slide, resolvedTheme, image, width, height);
            }
            return FinalizeDesignerAccessibility(slide, title);
        }

        /// <summary>Adds an editable architecture story using native shapes and connectors.</summary>
        public static PowerPointSlide AddDesignerArchitectureSlide(this PowerPointPresentation presentation,
            string title, string? subtitle, PowerPointArchitectureContent content,
            PowerPointDesignTheme? theme = null, PowerPointArchitectureSlideOptions? options = null) {
            if (presentation == null) throw new ArgumentNullException(nameof(presentation));
            if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("Title cannot be empty.", nameof(title));
            if (content == null) throw new ArgumentNullException(nameof(content));
            if (content.Nodes.Count > 12) {
                throw new ArgumentOutOfRangeException(nameof(content),
                    "Architecture slides support up to 12 nodes; split larger systems into focused views.");
            }
            PowerPointDesignTheme resolvedTheme = ResolveTheme(theme);
            PowerPointArchitectureSlideOptions resolved = options ?? new PowerPointArchitectureSlideOptions();
            PowerPointSlide slide = AddDesignerSlide(presentation, resolved);
            double width = presentation.SlideSize.WidthCm;
            double height = presentation.SlideSize.HeightCm;
            PrepareLightStorySlide(slide, resolvedTheme, resolved, title, subtitle, width, height);
            PowerPointArchitectureLayoutVariant variant = ResolveArchitectureVariant(resolved, content);
            if (variant == PowerPointArchitectureLayoutVariant.HubSpoke) {
                AddArchitectureHubSpoke(slide, resolvedTheme, content, width, height);
            } else {
                AddArchitectureLayers(slide, resolvedTheme, content, width, height);
            }
            return FinalizeDesignerAccessibility(slide, title);
        }

        internal static PowerPointScreenshotStoryLayoutVariant ResolveScreenshotVariant(
            PowerPointScreenshotStorySlideOptions options, PowerPointImageAsset image,
            IReadOnlyCollection<string> narrative) {
            if (options.Variant != PowerPointScreenshotStoryLayoutVariant.Auto) return options.Variant;
            if (narrative.Count > 0 || options.DesignIntent.LayoutStrategy == PowerPointAutoLayoutStrategy.Compact)
                return PowerPointScreenshotStoryLayoutVariant.SplitNarrative;
            return image.Annotations.Count > 0
                ? PowerPointScreenshotStoryLayoutVariant.HeroAnnotated
                : PowerPointScreenshotStoryLayoutVariant.SplitNarrative;
        }

        internal static PowerPointArchitectureLayoutVariant ResolveArchitectureVariant(
            PowerPointArchitectureSlideOptions options, PowerPointArchitectureContent content) {
            if (options.Variant != PowerPointArchitectureLayoutVariant.Auto) return options.Variant;
            return content.Nodes.Any(node => !string.IsNullOrWhiteSpace(node.Group))
                ? PowerPointArchitectureLayoutVariant.Layered
                : PowerPointArchitectureLayoutVariant.HubSpoke;
        }

        private static void AddScreenshotHero(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointImageAsset image, double width, double height) {
            PowerPointLayoutBox bounds = PowerPointLayoutBox.FromCentimeters(1.5, 3.55, width - 3, height - 5.5);
            PowerPointAutoShape frame = slide.AddRectangleCm(bounds.LeftCm - 0.08, bounds.TopCm - 0.08,
                bounds.WidthCm + 0.16, bounds.HeightCm + 0.16, "Screenshot Hero Frame");
            frame.FillColor = theme.AccentDarkColor;
            frame.OutlineColor = theme.AccentDarkColor;
            frame.SetShadow("000000", blurPoints: 5, distancePoints: 1.5, angleDegrees: 90, transparencyPercent: 82);
            slide.AddPicture(image, bounds);
            AddImageAnnotations(slide, theme, image, bounds, compact: true);
            AddStoryCaption(slide, theme, image.Caption, image.Provenance, 1.55, height - 1.65, width - 3.1);
        }

        private static void AddScreenshotSplitNarrative(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointImageAsset image, IReadOnlyList<string> narrative, double width, double height) {
            PowerPointLayoutBox body = PowerPointLayoutBox.FromCentimeters(1.5, 3.55, width - 3, height - 5.25);
            double imageWidth = body.WidthCm * 0.69;
            PowerPointLayoutBox imageBounds = PowerPointLayoutBox.FromCentimeters(body.LeftCm, body.TopCm,
                imageWidth, body.HeightCm);
            slide.AddPicture(image, imageBounds);
            AddImageAnnotations(slide, theme, image, imageBounds, compact: false);

            double railLeft = imageBounds.RightCm + 0.7;
            double railWidth = body.RightCm - railLeft;
            PowerPointAutoShape rail = slide.AddRectangleCm(railLeft, body.TopCm, railWidth, body.HeightCm,
                "Screenshot Narrative Rail");
            rail.FillColor = theme.SurfaceColor;
            rail.OutlineColor = theme.PanelBorderColor;
            AddText(slide, "What matters", railLeft + 0.45, body.TopCm + 0.45, railWidth - 0.9,
                0.55, 14, theme.AccentDarkColor, theme.HeadingFontName, bold: true);
            IReadOnlyList<string> resolved = narrative.Count > 0
                ? narrative
                : image.Annotations.Select(annotation => annotation.Detail ?? annotation.Label).ToList();
            AddText(slide, string.Join("\n", resolved.Select(point => "• " + point)), railLeft + 0.45,
                body.TopCm + 1.25, railWidth - 0.9, body.HeightCm - 1.65, 10,
                theme.SecondaryTextColor, theme.BodyFontName);
            AddStoryCaption(slide, theme, image.Caption, image.Provenance, 1.55, height - 1.62, width - 3.1);
        }

        private static void AddImageAnnotations(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointImageAsset image, PowerPointLayoutBox bounds, bool compact) {
            for (int index = 0; index < image.Annotations.Count; index++) {
                PowerPointImageAnnotation annotation = image.Annotations[index];
                string accent = annotation.Color ?? GetAccent(theme, index);
                double x = bounds.LeftCm + annotation.X * bounds.WidthCm;
                double y = bounds.TopCm + annotation.Y * bounds.HeightCm;
                double markerSize = compact ? 0.48 : 0.4;
                PowerPointAutoShape marker = slide.AddEllipseCm(x - markerSize / 2D, y - markerSize / 2D,
                    markerSize, markerSize, "Screenshot Annotation " + (index + 1));
                marker.FillColor = accent;
                marker.OutlineColor = theme.AccentContrastColor;
                marker.OutlineWidthPoints = 1.2;
                if (!compact) continue;
                double labelWidth = Math.Min(5.0, bounds.WidthCm * 0.26);
                double labelLeft = x + labelWidth + 0.4 <= bounds.RightCm
                    ? x + 0.38
                    : x - labelWidth - 0.38;
                double labelTop = Math.Max(bounds.TopCm + 0.15,
                    Math.Min(bounds.BottomCm - 0.95, y - 0.4));
                PowerPointAutoShape label = slide.AddRectangleCm(labelLeft, labelTop, labelWidth, 0.82,
                    "Screenshot Annotation Label " + (index + 1));
                label.FillColor = theme.AccentDarkColor;
                label.OutlineColor = accent;
                AddText(slide, annotation.Label, labelLeft + 0.2, labelTop + 0.15,
                    labelWidth - 0.4, 0.48, 9, theme.AccentContrastColor, theme.BodyFontName, bold: true);
            }
        }

        private static void AddArchitectureLayers(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointArchitectureContent content, double width, double height) {
            PowerPointLayoutBox body = PowerPointLayoutBox.FromCentimeters(1.5, 3.55, width - 3, height - 5.15);
            List<IGrouping<string, PowerPointArchitectureNode>> groups = content.Nodes
                .GroupBy(node => string.IsNullOrWhiteSpace(node.Group) ? "System" : node.Group!)
                .ToList();
            PowerPointLayoutBox[] rows = body.SplitRowsCm(groups.Count, 0.35);
            var bounds = new Dictionary<string, PowerPointLayoutBox>(StringComparer.OrdinalIgnoreCase);
            for (int row = 0; row < groups.Count; row++) {
                List<PowerPointArchitectureNode> nodes = groups[row].ToList();
                PowerPointLayoutBox groupBounds = rows[row];
                PowerPointAutoShape band = slide.AddRectangleCm(groupBounds.LeftCm, groupBounds.TopCm,
                    groupBounds.WidthCm, groupBounds.HeightCm, "Architecture Layer " + groups[row].Key);
                band.FillColor = row % 2 == 0 ? theme.SurfaceColor : theme.PanelColor;
                band.OutlineColor = theme.PanelBorderColor;
                AddText(slide, groups[row].Key.ToUpperInvariant(), groupBounds.LeftCm + 0.25,
                    groupBounds.TopCm + 0.18, 2.25, 0.4, 8, theme.MutedTextColor, theme.BodyFontName, bold: true);
                double nodeHeight = Math.Min(2.05, Math.Max(1.2, groupBounds.HeightCm - 0.9));
                PowerPointLayoutBox nodeArea = PowerPointLayoutBox.FromCentimeters(groupBounds.LeftCm + 2.5,
                    groupBounds.TopCm + (groupBounds.HeightCm - nodeHeight) / 2D,
                    groupBounds.WidthCm - 2.75, nodeHeight);
                PowerPointLayoutBox[] columns = nodeArea.SplitColumnsCm(nodes.Count, 0.35);
                for (int column = 0; column < nodes.Count; column++) bounds[nodes[column].Id] = columns[column];
            }
            AddArchitectureEdges(slide, theme, content.Edges, bounds);
            foreach (PowerPointArchitectureNode node in content.Nodes) AddArchitectureNode(slide, theme, node, bounds[node.Id]);
        }

        private static void AddArchitectureHubSpoke(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointArchitectureContent content, double width, double height) {
            PowerPointLayoutBox body = PowerPointLayoutBox.FromCentimeters(1.5, 3.55, width - 3, height - 5.15);
            var bounds = new Dictionary<string, PowerPointLayoutBox>(StringComparer.OrdinalIgnoreCase);
            PowerPointArchitectureNode hub = ResolveHub(content);
            bounds[hub.Id] = PowerPointLayoutBox.FromCentimeters(CenterX(body) - 2.35, CenterY(body) - 0.8,
                4.7, 1.6);
            List<PowerPointArchitectureNode> spokes = content.Nodes.Where(node => node != hub).ToList();
            double radiusX = Math.Max(3.1, body.WidthCm * 0.39);
            double radiusY = Math.Max(2.0, body.HeightCm * 0.36);
            for (int index = 0; index < spokes.Count; index++) {
                double angle = -Math.PI / 2D + 2D * Math.PI * index / Math.Max(1, spokes.Count);
                double centerX = CenterX(body) + Math.Cos(angle) * radiusX;
                double centerY = CenterY(body) + Math.Sin(angle) * radiusY;
                bounds[spokes[index].Id] = PowerPointLayoutBox.FromCentimeters(centerX - 1.85,
                    centerY - 0.68, 3.7, 1.36);
            }
            AddArchitectureEdges(slide, theme, content.Edges, bounds);
            foreach (PowerPointArchitectureNode node in content.Nodes) AddArchitectureNode(slide, theme, node, bounds[node.Id]);
        }

        private static PowerPointArchitectureNode ResolveHub(PowerPointArchitectureContent content) {
            return content.Nodes
                .OrderByDescending(node => content.Edges.Count(edge =>
                    string.Equals(edge.FromId, node.Id, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(edge.ToId, node.Id, StringComparison.OrdinalIgnoreCase)))
                .First();
        }

        private static void AddArchitectureEdges(PowerPointSlide slide, PowerPointDesignTheme theme,
            IReadOnlyList<PowerPointArchitectureEdge> edges,
            IReadOnlyDictionary<string, PowerPointLayoutBox> bounds) {
            for (int index = 0; index < edges.Count; index++) {
                PowerPointArchitectureEdge edge = edges[index];
                PowerPointLayoutBox from = bounds[edge.FromId];
                PowerPointLayoutBox to = bounds[edge.ToId];
                PowerPointAutoShape connector = slide.AddLineCm(CenterX(from), CenterY(from),
                    CenterX(to), CenterY(to), "Architecture Edge " + (index + 1));
                connector.OutlineColor = GetAccent(theme, index);
                connector.OutlineWidthPoints = 1.25;
                connector.SetLineEnds(null, A.LineEndValues.Triangle, A.LineEndWidthValues.Small,
                    A.LineEndLengthValues.Small);
                if (!string.IsNullOrWhiteSpace(edge.Label)) {
                    double labelX = (CenterX(from) + CenterX(to)) / 2D - 1.0;
                    double labelY = (CenterY(from) + CenterY(to)) / 2D - 0.22;
                    AddText(slide, edge.Label!, labelX, labelY, 2.0, 0.46, 8,
                        theme.MutedTextColor, theme.BodyFontName, bold: true);
                }
            }
        }

        private static void AddArchitectureNode(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointArchitectureNode node, PowerPointLayoutBox bounds) {
            PowerPointAutoShape panel = slide.AddShapeCm(A.ShapeTypeValues.RoundRectangle, bounds.LeftCm,
                bounds.TopCm, bounds.WidthCm, bounds.HeightCm, "Architecture Node " + node.Id);
            panel.FillColor = theme.BackgroundColor;
            panel.OutlineColor = theme.AccentColor;
            panel.OutlineWidthPoints = 1.4;
            panel.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 88);
            AddText(slide, node.Title, bounds.LeftCm + 0.25, bounds.TopCm + 0.2,
                bounds.WidthCm - 0.5, string.IsNullOrWhiteSpace(node.Body) ? bounds.HeightCm - 0.4 : 0.42,
                10, theme.PrimaryTextColor, theme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(node.Body)) {
                AddText(slide, node.Body!, bounds.LeftCm + 0.25, bounds.TopCm + 0.7,
                    bounds.WidthCm - 0.5, Math.Max(0.35, bounds.HeightCm - 0.9), 8,
                    theme.SecondaryTextColor, theme.BodyFontName);
            }
        }

        private static double CenterX(PowerPointLayoutBox bounds) => bounds.LeftCm + bounds.WidthCm / 2D;

        private static double CenterY(PowerPointLayoutBox bounds) => bounds.TopCm + bounds.HeightCm / 2D;
    }
}
