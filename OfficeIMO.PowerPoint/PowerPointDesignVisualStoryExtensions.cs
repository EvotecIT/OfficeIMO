using System;
using System.Collections.Generic;
using System.Linq;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointDesignExtensions {
        /// <summary>Adds a screenshot story with semantic crop, alternative text, caption, provenance, and annotations.</summary>
        public static PowerPointSlide AddDesignerScreenshotStorySlide(this PowerPointDocument presentation,
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
            double width = presentation.SlideSize.Width.ToCentimeters();
            double height = presentation.SlideSize.Height.ToCentimeters();
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
        public static PowerPointSlide AddDesignerArchitectureSlide(this PowerPointDocument presentation,
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
            double width = presentation.SlideSize.Width.ToCentimeters();
            double height = presentation.SlideSize.Height.ToCentimeters();
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
            PowerPointLayoutBox bounds = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.55),OfficeIMO.OfficeLength.Centimeters(width - 3),OfficeIMO.OfficeLength.Centimeters(height - 5.5));
            PowerPointAutoShape frame = slide.AddRectangle(OfficeIMO.OfficeLength.Centimeters(bounds.Left.ToCentimeters() - 0.08), OfficeIMO.OfficeLength.Centimeters(bounds.Top.ToCentimeters() - 0.08),
                OfficeIMO.OfficeLength.Centimeters(bounds.Width.ToCentimeters() + 0.16), OfficeIMO.OfficeLength.Centimeters(bounds.Height.ToCentimeters() + 0.16), "Screenshot Hero Frame");
            frame.FillColor = theme.AccentDarkColor;
            frame.OutlineColor = theme.AccentDarkColor;
            frame.SetShadow("000000", blurPoints: 5, distancePoints: 1.5, angleDegrees: 90, transparencyPercent: 82);
            PowerPointPicture picture = slide.AddImage(image,bounds);
            AddImageAnnotations(slide, theme, image, picture, compact: true);
            AddStoryCaption(slide, theme, image.Caption, image.Provenance, 1.55, height - 1.65, width - 3.1);
        }

        private static void AddScreenshotSplitNarrative(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointImageAsset image, IReadOnlyList<string> narrative, double width, double height) {
            PowerPointLayoutBox body = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.55),OfficeIMO.OfficeLength.Centimeters(width - 3),OfficeIMO.OfficeLength.Centimeters(height - 5.25));
            double imageWidth = body.Width.ToCentimeters() * 0.69;
            PowerPointLayoutBox imageBounds = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(body.Left.ToCentimeters()),OfficeIMO.OfficeLength.Centimeters(body.Top.ToCentimeters()),                OfficeIMO.OfficeLength.Centimeters(imageWidth),OfficeIMO.OfficeLength.Centimeters(body.Height.ToCentimeters()));
            PowerPointPicture picture = slide.AddImage(image,imageBounds);
            AddImageAnnotations(slide, theme, image, picture, compact: false);

            double railLeft = imageBounds.Right.ToCentimeters() + 0.7;
            double railWidth = body.Right.ToCentimeters() - railLeft;
            PowerPointAutoShape rail = slide.AddRectangle(OfficeIMO.OfficeLength.Centimeters(railLeft), body.Top, OfficeIMO.OfficeLength.Centimeters(railWidth), body.Height,
                "Screenshot Narrative Rail");
            rail.FillColor = theme.SurfaceColor;
            rail.OutlineColor = theme.PanelBorderColor;
            AddText(slide, "What matters", railLeft + 0.45, body.Top.ToCentimeters() + 0.45, railWidth - 0.9,
                0.55, 14, theme.AccentDarkColor, theme.HeadingFontName, bold: true);
            IReadOnlyList<string> resolved = narrative.Count > 0
                ? narrative
                : image.Annotations.Select(annotation => annotation.Detail ?? annotation.Label).ToList();
            AddText(slide, string.Join("\n", resolved.Select(point => "• " + point)), railLeft + 0.45,
                body.Top.ToCentimeters() + 1.25, railWidth - 0.9, body.Height.ToCentimeters() - 1.65, 10,
                theme.SecondaryTextColor, theme.BodyFontName);
            AddStoryCaption(slide, theme, image.Caption, image.Provenance, 1.55, height - 1.62, width - 3.1);
        }

        private static void AddImageAnnotations(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointImageAsset image, PowerPointPicture picture, bool compact) {
            for (int index = 0; index < image.Annotations.Count; index++) {
                PowerPointImageAnnotation annotation = image.Annotations[index];
                if (!TryProjectImageAnnotation(picture, annotation, out double x, out double y)) continue;
                string accent = annotation.Color ?? GetAccent(theme, index);
                double markerSize = compact ? 0.48 : 0.4;
                PowerPointAutoShape marker = slide.AddEllipse(OfficeIMO.OfficeLength.Centimeters(x - markerSize / 2D), OfficeIMO.OfficeLength.Centimeters(y - markerSize / 2D),
                    OfficeIMO.OfficeLength.Centimeters(markerSize), OfficeIMO.OfficeLength.Centimeters(markerSize), "Screenshot Annotation " + (index + 1));
                marker.FillColor = accent;
                marker.OutlineColor = theme.AccentContrastColor;
                marker.OutlineWidthPoints = 1.2;
                if (!compact) continue;
                double labelWidth = Math.Min(5.0, picture.Width.ToCentimeters() * 0.26);
                double labelLeft = x + labelWidth + 0.4 <= picture.Right.ToCentimeters()
                    ? x + 0.38
                    : x - labelWidth - 0.38;
                double labelTop = Math.Max(picture.Top.ToCentimeters() + 0.15,
                    Math.Min(picture.Bottom.ToCentimeters() - 0.95, y - 0.4));
                PowerPointAutoShape label = slide.AddRectangle(OfficeIMO.OfficeLength.Centimeters(labelLeft), OfficeIMO.OfficeLength.Centimeters(labelTop), OfficeIMO.OfficeLength.Centimeters(labelWidth), OfficeIMO.OfficeLength.Centimeters(0.82),
                    "Screenshot Annotation Label " + (index + 1));
                label.FillColor = theme.AccentDarkColor;
                label.OutlineColor = accent;
                AddText(slide, annotation.Label, labelLeft + 0.2, labelTop + 0.15,
                    labelWidth - 0.4, 0.48, 9, theme.AccentContrastColor, theme.BodyFontName, bold: true);
            }
        }

        private static bool TryProjectImageAnnotation(PowerPointPicture picture,
            PowerPointImageAnnotation annotation, out double x, out double y) {
            double left = picture.CropLeftRatio;
            double top = picture.CropTopRatio;
            double right = 1D - picture.CropRightRatio;
            double bottom = 1D - picture.CropBottomRatio;
            if (annotation.X < left || annotation.X > right ||
                annotation.Y < top || annotation.Y > bottom ||
                right <= left || bottom <= top) {
                x = 0D;
                y = 0D;
                return false;
            }

            x = picture.Left.ToCentimeters() + (annotation.X - left) / (right - left) * picture.Width.ToCentimeters();
            y = picture.Top.ToCentimeters() + (annotation.Y - top) / (bottom - top) * picture.Height.ToCentimeters();
            return true;
        }

        private static void AddArchitectureLayers(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointArchitectureContent content, double width, double height) {
            PowerPointLayoutBox body = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.55),OfficeIMO.OfficeLength.Centimeters(width - 3),OfficeIMO.OfficeLength.Centimeters(height - 5.15));
            List<IGrouping<string, PowerPointArchitectureNode>> groups = content.Nodes
                .GroupBy(node => string.IsNullOrWhiteSpace(node.Group) ? "System" : node.Group!)
                .ToList();
            PowerPointLayoutBox[] rows = body.SplitRows(groups.Count,OfficeIMO.OfficeLength.Centimeters(0.35));
            var bounds = new Dictionary<string, PowerPointLayoutBox>(StringComparer.OrdinalIgnoreCase);
            for (int row = 0; row < groups.Count; row++) {
                List<PowerPointArchitectureNode> nodes = groups[row].ToList();
                PowerPointLayoutBox groupBounds = rows[row];
                PowerPointAutoShape band = slide.AddRectangle(groupBounds.Left, groupBounds.Top,
                    groupBounds.Width, groupBounds.Height, "Architecture Layer " + groups[row].Key);
                band.FillColor = row % 2 == 0 ? theme.SurfaceColor : theme.PanelColor;
                band.OutlineColor = theme.PanelBorderColor;
                AddText(slide, groups[row].Key.ToUpperInvariant(), groupBounds.Left.ToCentimeters() + 0.25,
                    groupBounds.Top.ToCentimeters() + 0.18, 2.25, 0.4, 8, theme.MutedTextColor, theme.BodyFontName, bold: true);
                double nodeHeight = Math.Min(2.05, Math.Max(1.2, groupBounds.Height.ToCentimeters() - 0.9));
                PowerPointLayoutBox nodeArea = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(groupBounds.Left.ToCentimeters() + 2.5),                    OfficeIMO.OfficeLength.Centimeters(groupBounds.Top.ToCentimeters() + (groupBounds.Height.ToCentimeters() - nodeHeight) / 2D),                    OfficeIMO.OfficeLength.Centimeters(groupBounds.Width.ToCentimeters() - 2.75),OfficeIMO.OfficeLength.Centimeters(nodeHeight));
                PowerPointLayoutBox[] columns = nodeArea.SplitColumns(nodes.Count,OfficeIMO.OfficeLength.Centimeters(0.35));
                for (int column = 0; column < nodes.Count; column++) bounds[nodes[column].Id] = columns[column];
            }
            AddArchitectureEdges(slide, theme, content.Edges, bounds);
            foreach (PowerPointArchitectureNode node in content.Nodes) AddArchitectureNode(slide, theme, node, bounds[node.Id]);
        }

        private static void AddArchitectureHubSpoke(PowerPointSlide slide, PowerPointDesignTheme theme,
            PowerPointArchitectureContent content, double width, double height) {
            PowerPointLayoutBox body = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(1.5),OfficeIMO.OfficeLength.Centimeters(3.55),OfficeIMO.OfficeLength.Centimeters(width - 3),OfficeIMO.OfficeLength.Centimeters(height - 5.15));
            var bounds = new Dictionary<string, PowerPointLayoutBox>(StringComparer.OrdinalIgnoreCase);
            PowerPointArchitectureNode hub = ResolveHub(content);
            bounds[hub.Id] = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(CenterX(body) - 2.35),OfficeIMO.OfficeLength.Centimeters(CenterY(body) - 0.8),                OfficeIMO.OfficeLength.Centimeters(4.7),OfficeIMO.OfficeLength.Centimeters(1.6));
            List<PowerPointArchitectureNode> spokes = content.Nodes.Where(node => node != hub).ToList();
            double radiusX = Math.Max(3.1, body.Width.ToCentimeters() * 0.39);
            double radiusY = Math.Max(2.0, body.Height.ToCentimeters() * 0.36);
            for (int index = 0; index < spokes.Count; index++) {
                double angle = -Math.PI / 2D + 2D * Math.PI * index / Math.Max(1, spokes.Count);
                double centerX = CenterX(body) + Math.Cos(angle) * radiusX;
                double centerY = CenterY(body) + Math.Sin(angle) * radiusY;
                bounds[spokes[index].Id] = new OfficeIMO.PowerPoint.PowerPointLayoutBox(OfficeIMO.OfficeLength.Centimeters(centerX - 1.85),                    OfficeIMO.OfficeLength.Centimeters(centerY - 0.68),OfficeIMO.OfficeLength.Centimeters(3.7),OfficeIMO.OfficeLength.Centimeters(1.36));
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
                PowerPointAutoShape connector = slide.AddLine(OfficeIMO.OfficeLength.Centimeters(CenterX(from)), OfficeIMO.OfficeLength.Centimeters(CenterY(from)),
                    OfficeIMO.OfficeLength.Centimeters(CenterX(to)), OfficeIMO.OfficeLength.Centimeters(CenterY(to)), "Architecture Edge " + (index + 1));
                connector.OutlineColor = GetAccent(theme, index);
                connector.OutlineWidthPoints = 1.25;
                connector.SetLineEnds(null, OfficeLineMarkerKind.Triangle, PowerPointLineEndWidth.Small,
                    PowerPointLineEndLength.Small);
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
            PowerPointAutoShape panel = slide.AddShape(OfficePresetShapeType.RoundRectangle, bounds.Left,
                bounds.Top, bounds.Width, bounds.Height, "Architecture Node " + node.Id);
            panel.FillColor = theme.BackgroundColor;
            panel.OutlineColor = theme.AccentColor;
            panel.OutlineWidthPoints = 1.4;
            panel.SetShadow("000000", blurPoints: 3, distancePoints: 0.8, angleDegrees: 90, transparencyPercent: 88);
            AddText(slide, node.Title, bounds.Left.ToCentimeters() + 0.25, bounds.Top.ToCentimeters() + 0.2,
                bounds.Width.ToCentimeters() - 0.5, string.IsNullOrWhiteSpace(node.Body) ? bounds.Height.ToCentimeters() - 0.4 : 0.42,
                10, theme.PrimaryTextColor, theme.HeadingFontName, bold: true);
            if (!string.IsNullOrWhiteSpace(node.Body)) {
                AddText(slide, node.Body!, bounds.Left.ToCentimeters() + 0.25, bounds.Top.ToCentimeters() + 0.7,
                    bounds.Width.ToCentimeters() - 0.5, Math.Max(0.35, bounds.Height.ToCentimeters() - 0.9), 8,
                    theme.SecondaryTextColor, theme.BodyFontName);
            }
        }

        private static double CenterX(PowerPointLayoutBox bounds) => bounds.Left.ToCentimeters() + bounds.Width.ToCentimeters() / 2D;

        private static double CenterY(PowerPointLayoutBox bounds) => bounds.Top.ToCentimeters() + bounds.Height.ToCentimeters() / 2D;
    }
}
