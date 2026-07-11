using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>Reads a template inventory from an existing presentation or PowerPoint template file.</summary>
        public static PowerPointTemplateInventory InspectTemplate(string templatePath) {
            if (string.IsNullOrWhiteSpace(templatePath)) {
                throw new ArgumentException("Template path cannot be empty.", nameof(templatePath));
            }
            if (!File.Exists(templatePath)) {
                throw new FileNotFoundException("PowerPoint template was not found.", templatePath);
            }

            using PowerPointPresentation presentation = OpenRead(templatePath);
            return presentation.InspectTemplate();
        }

        /// <summary>Inventories masters, layouts, placeholders, theme tokens, assets, footer content, and safe areas.</summary>
        public PowerPointTemplateInventory InspectTemplate() {
            ThrowIfDisposed();
            var masters = new List<PowerPointTemplateMasterInfo>();
            var assets = new List<PowerPointTemplateAssetInfo>();
            var footerContents = new List<string>();
            SlideMasterPart[] masterParts = _presentationPart.SlideMasterParts.ToArray();
            PowerPointLayoutBox slideBounds = new PowerPointLayoutBox(0L, 0L, SlideSize.WidthEmus,
                SlideSize.HeightEmus);
            PowerPointLayoutBox fallbackSafeArea = SlideSize.GetContentBoxCm(1D);

            for (int masterIndex = 0; masterIndex < masterParts.Length; masterIndex++) {
                SlideMasterPart masterPart = masterParts[masterIndex];
                string masterName = masterPart.SlideMaster?.CommonSlideData?.Name?.Value
                    ?? "Master " + (masterIndex + 1);
                string themeName = masterPart.ThemePart?.Theme?.Name?.Value ?? string.Empty;
                var layouts = new List<PowerPointTemplateLayoutInfo>();
                SlideLayoutPart[] layoutParts = masterPart.SlideLayoutParts.ToArray();

                CollectAssets(masterPart, masterPart.SlideMaster?.CommonSlideData?.ShapeTree, masterIndex, null,
                    assets);
                for (int layoutIndex = 0; layoutIndex < layoutParts.Length; layoutIndex++) {
                    SlideLayoutPart layoutPart = layoutParts[layoutIndex];
                    SlideLayout? layout = layoutPart.SlideLayout;
                    ShapeTree? tree = layout?.CommonSlideData?.ShapeTree;
                    var placeholders = new List<PowerPointTemplatePlaceholderInfo>();
                    if (tree != null) {
                        foreach (OpenXmlElement element in tree.ChildElements) {
                            PlaceholderShape? placeholder = GetTemplatePlaceholder(element);
                            if (placeholder == null) continue;
                            string name = GetTemplateElementName(element);
                            string? defaultText = string.IsNullOrWhiteSpace(element.InnerText)
                                ? null
                                : element.InnerText;
                            PowerPointTemplatePlaceholderRole role = InferPlaceholderRole(
                                placeholder.Type?.Value, name);
                            placeholders.Add(new PowerPointTemplatePlaceholderInfo(name,
                                placeholder.Type?.Value, placeholder.Index?.Value, role,
                                GetTemplateElementBounds(element), defaultText));
                            if (role == PowerPointTemplatePlaceholderRole.Footer && defaultText != null) {
                                footerContents.Add(defaultText);
                            }
                        }
                    }

                    PowerPointLayoutBox safeArea = ResolveSafeArea(placeholders, fallbackSafeArea, slideBounds);
                    PowerPointLayoutBox? titleArea = placeholders.FirstOrDefault(placeholder =>
                        placeholder.Role == PowerPointTemplatePlaceholderRole.Title)?.Bounds;
                    layouts.Add(new PowerPointTemplateLayoutInfo(masterIndex, layoutIndex,
                        layout?.CommonSlideData?.Name?.Value ?? string.Empty, layout?.Type?.Value,
                        placeholders, safeArea, titleArea));
                    CollectAssets(layoutPart, tree, masterIndex, layoutIndex, assets);
                }

                masters.Add(new PowerPointTemplateMasterInfo(masterIndex, masterName, themeName,
                    GetThemeColors(masterIndex).ToDictionary(pair => pair.Key, pair => pair.Value),
                    GetThemeFonts(masterIndex), layouts));
            }

            string? sourcePath = string.IsNullOrWhiteSpace(_filePath) ? null : Path.GetFullPath(_filePath);
            return new PowerPointTemplateInventory(sourcePath, Slides.Count, slideBounds, masters, assets,
                footerContents);
        }

        private static void CollectAssets(OpenXmlPartContainer owner, ShapeTree? tree, int masterIndex,
            int? layoutIndex, IList<PowerPointTemplateAssetInfo> assets) {
            if (tree == null) return;
            foreach (DocumentFormat.OpenXml.Presentation.Picture picture in
                     tree.Descendants<DocumentFormat.OpenXml.Presentation.Picture>()) {
                NonVisualDrawingProperties? properties = picture.NonVisualPictureProperties?
                    .NonVisualDrawingProperties;
                string name = properties?.Name?.Value ?? string.Empty;
                string? description = properties?.Description?.Value;
                string combined = name + " " + description;
                PowerPointTemplateAssetKind kind = ContainsAny(combined, "logo", "brand", "wordmark")
                    ? PowerPointTemplateAssetKind.Logo
                    : PowerPointTemplateAssetKind.Picture;
                string? contentType = ResolvePictureContentType(owner, picture);
                assets.Add(new PowerPointTemplateAssetInfo(kind, masterIndex, layoutIndex, name, description,
                    contentType, GetTemplateElementBounds(picture)));
            }
        }

        private static string? ResolvePictureContentType(OpenXmlPartContainer owner,
            DocumentFormat.OpenXml.Presentation.Picture picture) {
            string? relationshipId = picture.BlipFill?.Blip?.Embed?.Value;
            if (string.IsNullOrWhiteSpace(relationshipId)) return null;
            try {
                return (owner.GetPartById(relationshipId!) as ImagePart)?.ContentType;
            } catch (ArgumentOutOfRangeException) {
                return null;
            }
        }

        private static PowerPointLayoutBox ResolveSafeArea(
            IList<PowerPointTemplatePlaceholderInfo> placeholders, PowerPointLayoutBox fallback,
            PowerPointLayoutBox slideBounds) {
            List<PowerPointLayoutBox> contentBounds = placeholders
                .Where(placeholder => placeholder.Bounds.HasValue &&
                    placeholder.Role != PowerPointTemplatePlaceholderRole.Footer &&
                    placeholder.Role != PowerPointTemplatePlaceholderRole.Date &&
                    placeholder.Role != PowerPointTemplatePlaceholderRole.SlideNumber &&
                    placeholder.Role != PowerPointTemplatePlaceholderRole.Title)
                .Select(placeholder => placeholder.Bounds!.Value)
                .ToList();
            if (contentBounds.Count == 0) return fallback;

            long left = Math.Max(slideBounds.Left, contentBounds.Min(bounds => bounds.Left));
            long top = Math.Max(slideBounds.Top, contentBounds.Min(bounds => bounds.Top));
            long right = Math.Min(slideBounds.Right, contentBounds.Max(bounds => bounds.Right));
            long bottom = Math.Min(slideBounds.Bottom, contentBounds.Max(bounds => bounds.Bottom));
            return right > left && bottom > top
                ? new PowerPointLayoutBox(left, top, right - left, bottom - top)
                : fallback;
        }

        private static PowerPointTemplatePlaceholderRole InferPlaceholderRole(PlaceholderValues? type,
            string name) {
            if (type.HasValue) {
                PlaceholderValues value = type.Value;
                if (value == PlaceholderValues.Title || value == PlaceholderValues.CenteredTitle)
                    return PowerPointTemplatePlaceholderRole.Title;
                if (value == PlaceholderValues.SubTitle)
                    return PowerPointTemplatePlaceholderRole.Subtitle;
                if (value == PlaceholderValues.Body)
                    return PowerPointTemplatePlaceholderRole.Body;
                if (value == PlaceholderValues.Picture || value == PlaceholderValues.Media ||
                    value == PlaceholderValues.ClipArt)
                    return PowerPointTemplatePlaceholderRole.Image;
                if (value == PlaceholderValues.Chart)
                    return PowerPointTemplatePlaceholderRole.Chart;
                if (value == PlaceholderValues.Table)
                    return PowerPointTemplatePlaceholderRole.Table;
                if (value == PlaceholderValues.Footer)
                    return PowerPointTemplatePlaceholderRole.Footer;
                if (value == PlaceholderValues.DateAndTime)
                    return PowerPointTemplatePlaceholderRole.Date;
                if (value == PlaceholderValues.SlideNumber)
                    return PowerPointTemplatePlaceholderRole.SlideNumber;
                if (value == PlaceholderValues.Object)
                    return PowerPointTemplatePlaceholderRole.Content;
            }

            if (ContainsAny(name, "title", "heading")) return PowerPointTemplatePlaceholderRole.Title;
            if (ContainsAny(name, "subtitle", "subheading")) return PowerPointTemplatePlaceholderRole.Subtitle;
            if (ContainsAny(name, "image", "picture", "photo", "screenshot", "logo"))
                return PowerPointTemplatePlaceholderRole.Image;
            if (ContainsAny(name, "chart", "graph")) return PowerPointTemplatePlaceholderRole.Chart;
            if (ContainsAny(name, "table", "grid")) return PowerPointTemplatePlaceholderRole.Table;
            if (ContainsAny(name, "body", "content", "copy")) return PowerPointTemplatePlaceholderRole.Body;
            if (ContainsAny(name, "footer")) return PowerPointTemplatePlaceholderRole.Footer;
            return PowerPointTemplatePlaceholderRole.Unknown;
        }

        private static bool ContainsAny(string? value, params string[] tokens) {
            string text = value ?? string.Empty;
            for (int index = 0; index < tokens.Length; index++) {
                if (text.IndexOf(tokens[index], StringComparison.OrdinalIgnoreCase) >= 0) return true;
            }
            return false;
        }

        private static PlaceholderShape? GetTemplatePlaceholder(OpenXmlElement element) => element switch {
            Shape shape => shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
            DocumentFormat.OpenXml.Presentation.Picture picture => picture.NonVisualPictureProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape,
            GraphicFrame frame => frame.NonVisualGraphicFrameProperties?
                .ApplicationNonVisualDrawingProperties?.PlaceholderShape,
            _ => null
        };

        private static string GetTemplateElementName(OpenXmlElement element) => element switch {
            Shape shape => shape.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
            DocumentFormat.OpenXml.Presentation.Picture picture => picture.NonVisualPictureProperties?
                .NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
            GraphicFrame frame => frame.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value
                ?? string.Empty,
            _ => string.Empty
        };

        private static PowerPointLayoutBox? GetTemplateElementBounds(OpenXmlElement element) => element switch {
            Shape shape => GetTemplateBounds(shape.ShapeProperties?.Transform2D),
            DocumentFormat.OpenXml.Presentation.Picture picture => GetTemplateBounds(
                picture.ShapeProperties?.Transform2D),
            GraphicFrame frame => GetTemplateBounds(frame.Transform),
            _ => null
        };

        private static PowerPointLayoutBox? GetTemplateBounds(A.Transform2D? transform) {
            long? x = transform?.Offset?.X?.Value;
            long? y = transform?.Offset?.Y?.Value;
            long? width = transform?.Extents?.Cx?.Value;
            long? height = transform?.Extents?.Cy?.Value;
            return x.HasValue && y.HasValue && width.HasValue && height.HasValue
                ? new PowerPointLayoutBox(x.Value, y.Value, width.Value, height.Value)
                : null;
        }

        private static PowerPointLayoutBox? GetTemplateBounds(Transform? transform) {
            long? x = transform?.Offset?.X?.Value;
            long? y = transform?.Offset?.Y?.Value;
            long? width = transform?.Extents?.Cx?.Value;
            long? height = transform?.Extents?.Cy?.Value;
            return x.HasValue && y.HasValue && width.HasValue && height.HasValue
                ? new PowerPointLayoutBox(x.Value, y.Value, width.Value, height.Value)
                : null;
        }
    }
}
