using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        ///     Gets the index of the layout used by this slide.
        /// </summary>
        public int LayoutIndex {
            get {
                SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
                if (layoutPart == null) {
                    return -1;
                }

                SlideMasterPart master = layoutPart.GetParentParts().OfType<SlideMasterPart>().First();
                SlideLayoutPart[] layouts = master.SlideLayoutParts.ToArray();
                for (int i = 0; i < layouts.Length; i++) {
                    if (layouts[i] == layoutPart) {
                        return i;
                    }
                }

                return -1;
            }
        }

        /// <summary>
        ///     Sets the slide layout using master and layout indexes.
        /// </summary>
        public void SetLayout(int masterIndex, int layoutIndex) {
            PresentationPart presentationPart = _slidePart.GetParentParts().OfType<PresentationPart>().First();

            SlideMasterPart[] masters = presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }

            SlideMasterPart masterPart = masters[masterIndex];
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            if (layoutIndex < 0 || layoutIndex >= layouts.Length) {
                throw new ArgumentOutOfRangeException(nameof(layoutIndex));
            }

            SlideLayoutPart layoutPart = layouts[layoutIndex];
            SlideLayoutPart? current = _slidePart.SlideLayoutPart;
            if (current != null) {
                string relId = _slidePart.GetIdOfPart(current);
                _slidePart.DeletePart(relId);
            }

            _slidePart.AddPart(layoutPart);
        }

        /// <summary>
        ///     Sets the slide layout using a layout type.
        /// </summary>
        public void SetLayout(SlideLayoutValues layoutType, int masterIndex = 0) {
            int layoutIndex = GetLayoutIndex(layoutType, masterIndex);
            SetLayout(masterIndex, layoutIndex);
        }

        /// <summary>
        ///     Sets the slide layout using a layout name.
        /// </summary>
        public void SetLayout(string layoutName, int masterIndex = 0, bool ignoreCase = true) {
            int layoutIndex = GetLayoutIndex(layoutName, masterIndex, ignoreCase);
            SetLayout(masterIndex, layoutIndex);
        }

        private int GetLayoutIndex(SlideLayoutValues layoutType, int masterIndex) {
            PresentationPart presentationPart = _slidePart.GetParentParts().OfType<PresentationPart>().First();
            SlideMasterPart[] masters = presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }

            SlideLayoutPart[] layouts = masters[masterIndex].SlideLayoutParts.ToArray();
            for (int i = 0; i < layouts.Length; i++) {
                SlideLayoutValues? type = layouts[i].SlideLayout?.Type?.Value;
                if (type == layoutType) {
                    return i;
                }
            }

            throw new InvalidOperationException($"Layout type '{layoutType}' not found for master {masterIndex}.");
        }

        private int GetLayoutIndex(string layoutName, int masterIndex, bool ignoreCase) {
            if (layoutName == null) {
                throw new ArgumentNullException(nameof(layoutName));
            }

            PresentationPart presentationPart = _slidePart.GetParentParts().OfType<PresentationPart>().First();
            SlideMasterPart[] masters = presentationPart.SlideMasterParts.ToArray();
            if (masterIndex < 0 || masterIndex >= masters.Length) {
                throw new ArgumentOutOfRangeException(nameof(masterIndex));
            }

            SlideLayoutPart[] layouts = masters[masterIndex].SlideLayoutParts.ToArray();
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            for (int i = 0; i < layouts.Length; i++) {
                string name = layouts[i].SlideLayout?.CommonSlideData?.Name?.Value ?? string.Empty;
                if (string.Equals(name, layoutName, comparison)) {
                    return i;
                }
            }

            throw new InvalidOperationException($"Layout '{layoutName}' not found for master {masterIndex}.");
        }

        /// <summary>
        ///     Textboxes that map to placeholders in the slide layout.
        /// </summary>
        public IReadOnlyList<PowerPointTextBox> Placeholders =>
            TextBoxes.Where(tb => tb.IsPlaceholder).ToList();

        /// <summary>
        ///     Retrieves the first placeholder textbox matching the specified type.
        /// </summary>
        public PowerPointTextBox? GetPlaceholder(PlaceholderValues placeholderType, uint? index = null) {
            IEnumerable<PowerPointTextBox> matches = TextBoxes
                .Where(tb => tb.PlaceholderType == placeholderType);

            if (index != null) {
                matches = matches.Where(tb => tb.PlaceholderIndex == index);
            }

            return matches.FirstOrDefault();
        }

        /// <summary>
        ///     Retrieves placeholders defined by the slide layout.
        /// </summary>
        public IReadOnlyList<PowerPointLayoutPlaceholderInfo> GetLayoutPlaceholders() {
            SlideLayoutPart? layoutPart = _slidePart.SlideLayoutPart;
            ShapeTree? shapeTree = layoutPart?.SlideLayout?.CommonSlideData?.ShapeTree;
            if (shapeTree == null) {
                return Array.Empty<PowerPointLayoutPlaceholderInfo>();
            }

            List<PowerPointLayoutPlaceholderInfo> placeholders = new();
            foreach (OpenXmlElement element in shapeTree.ChildElements) {
                PlaceholderShape? placeholder = GetLayoutPlaceholderShape(element);
                if (placeholder == null) {
                    continue;
                }

                string name = GetLayoutElementName(element);
                PowerPointLayoutBox? bounds = GetLayoutElementBounds(element);
                placeholders.Add(new PowerPointLayoutPlaceholderInfo(
                    name,
                    placeholder.Type?.Value,
                    placeholder.Index?.Value,
                    bounds));
            }

            return placeholders;
        }

        /// <summary>
        ///     Retrieves a layout placeholder by type and optional index.
        /// </summary>
        public PowerPointLayoutPlaceholderInfo? GetLayoutPlaceholder(PlaceholderValues placeholderType, uint? index = null) {
            foreach (PowerPointLayoutPlaceholderInfo placeholder in GetLayoutPlaceholders()) {
                if (placeholder.PlaceholderType != placeholderType) {
                    continue;
                }

                if (index != null && placeholder.PlaceholderIndex != index) {
                    continue;
                }

                return placeholder;
            }

            return null;
        }

        /// <summary>
        ///     Retrieves layout placeholder bounds by type and optional index.
        /// </summary>
        public PowerPointLayoutBox? GetLayoutPlaceholderBounds(PlaceholderValues placeholderType, uint? index = null) {
            PowerPointLayoutPlaceholderInfo? placeholder = GetLayoutPlaceholder(placeholderType, index);
            return placeholder?.Bounds;
        }

        private static PlaceholderShape? GetLayoutPlaceholderShape(OpenXmlElement element) {
            return element switch {
                Shape s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                DocumentFormat.OpenXml.Presentation.Picture p => p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                GraphicFrame g => g.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                _ => null
            };
        }

        private static string GetLayoutElementName(OpenXmlElement element) {
            return element switch {
                Shape s => s.NonVisualShapeProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
                DocumentFormat.OpenXml.Presentation.Picture p => p.NonVisualPictureProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
                GraphicFrame g => g.NonVisualGraphicFrameProperties?.NonVisualDrawingProperties?.Name?.Value ?? string.Empty,
                _ => string.Empty
            };
        }

        private static PowerPointLayoutBox? GetLayoutElementBounds(OpenXmlElement element) {
            return element switch {
                Shape s => GetLayoutElementBounds(s.ShapeProperties?.Transform2D),
                DocumentFormat.OpenXml.Presentation.Picture p => GetLayoutElementBounds(p.ShapeProperties?.Transform2D),
                GraphicFrame g => GetLayoutElementBounds(g.Transform),
                _ => null
            };
        }

        private static PowerPointLayoutBox? GetLayoutElementBounds(A.Transform2D? transform) {
            long? x = transform?.Offset?.X?.Value;
            long? y = transform?.Offset?.Y?.Value;
            long? cx = transform?.Extents?.Cx?.Value;
            long? cy = transform?.Extents?.Cy?.Value;
            if (x == null || y == null || cx == null || cy == null) {
                return null;
            }

            return new PowerPointLayoutBox(x.Value, y.Value, cx.Value, cy.Value);
        }

        private static PowerPointLayoutBox? GetLayoutElementBounds(Transform? transform) {
            long? x = transform?.Offset?.X?.Value;
            long? y = transform?.Offset?.Y?.Value;
            long? cx = transform?.Extents?.Cx?.Value;
            long? cy = transform?.Extents?.Cy?.Value;
            if (x == null || y == null || cx == null || cy == null) {
                return null;
            }

            return new PowerPointLayoutBox(x.Value, y.Value, cx.Value, cy.Value);
        }
    }
}
