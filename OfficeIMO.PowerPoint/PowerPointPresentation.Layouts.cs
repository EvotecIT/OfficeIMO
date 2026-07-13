using System;
using System.IO;
using System.Reflection;
using System.Runtime.ExceptionServices;
using System.Xml.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Drawing.Internal;
using A = DocumentFormat.OpenXml.Drawing;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointPresentation {
        /// <summary>
        ///     Returns the layouts available for a slide master.
        /// </summary>
        public IReadOnlyList<PowerPointSlideLayoutInfo> GetSlideLayouts(int masterIndex = 0) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            var results = new List<PowerPointSlideLayoutInfo>(layouts.Length);

            for (int i = 0; i < layouts.Length; i++) {
                SlideLayoutPart layoutPart = layouts[i];
                SlideLayout? layout = layoutPart.SlideLayout;
                string name = layout?.CommonSlideData?.Name?.Value ?? string.Empty;
                SlideLayoutValues? type = layout?.Type?.Value;
                string? relationshipId = masterPart.GetIdOfPart(layoutPart);
                results.Add(new PowerPointSlideLayoutInfo(masterIndex, i, name, type, relationshipId));
            }

            return results;
        }

        /// <summary>
        ///     Finds a layout index by layout type.
        /// </summary>
        public int GetLayoutIndex(SlideLayoutValues layoutType, int masterIndex = 0) {
            ThrowIfDisposed();
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            for (int i = 0; i < layouts.Length; i++) {
                SlideLayoutValues? type = layouts[i].SlideLayout?.Type?.Value;
                if (type == layoutType) {
                    return i;
                }
            }

            throw new InvalidOperationException($"Layout type '{layoutType}' not found for master {masterIndex}.");
        }

        /// <summary>
        ///     Finds a layout index by layout name.
        /// </summary>
        public int GetLayoutIndex(string layoutName, int masterIndex = 0, bool ignoreCase = true) {
            ThrowIfDisposed();
            if (layoutName == null) {
                throw new ArgumentNullException(nameof(layoutName));
            }

            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
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
        ///     Adds a slide using a layout type.
        /// </summary>
        public PowerPointSlide AddSlide(SlideLayoutValues layoutType, int masterIndex = 0) {
            int layoutIndex = GetLayoutIndex(layoutType, masterIndex);
            return AddSlide(masterIndex, layoutIndex);
        }

        /// <summary>
        ///     Adds a slide using a layout name.
        /// </summary>
        public PowerPointSlide AddSlide(string layoutName, int masterIndex = 0, bool ignoreCase = true) {
            int layoutIndex = GetLayoutIndex(layoutName, masterIndex, ignoreCase);
            return AddSlide(masterIndex, layoutIndex);
        }

        private SlideLayoutPart? FindMatchingLayout(SlideLayoutPart sourceLayoutPart) {
            SlideLayout? sourceLayout = sourceLayoutPart.SlideLayout;
            if (sourceLayout == null) {
                return null;
            }

            string? sourceName = sourceLayout.CommonSlideData?.Name?.Value;
            SlideLayoutValues? sourceType = sourceLayout.Type?.Value;

            foreach (SlideMasterPart masterPart in _presentationPart.SlideMasterParts) {
                foreach (SlideLayoutPart layoutPart in masterPart.SlideLayoutParts) {
                    SlideLayout? candidateLayout = layoutPart.SlideLayout;
                    if (candidateLayout == null) {
                        continue;
                    }

                    SlideLayoutValues? candidateType = candidateLayout.Type?.Value;
                    if (sourceType.HasValue && candidateType != sourceType) {
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(sourceName)) {
                        string? candidateName = candidateLayout.CommonSlideData?.Name?.Value;
                        if (!string.Equals(sourceName, candidateName, StringComparison.OrdinalIgnoreCase)) {
                            continue;
                        }
                    }

                    return layoutPart;
                }
            }

            return null;
        }

        private SlideLayoutPart GetSlideLayoutPart(int masterIndex, int layoutIndex) {
            SlideMasterPart masterPart = GetSlideMasterPart(masterIndex);
            SlideLayoutPart[] layouts = masterPart.SlideLayoutParts.ToArray();
            if (layoutIndex < 0 || layoutIndex >= layouts.Length) {
                throw new ArgumentOutOfRangeException(nameof(layoutIndex));
            }
            return layouts[layoutIndex];
        }

        /// <summary>
        ///     Retrieves a layout placeholder textbox for a master/layout pair.
        /// </summary>
        public PowerPointTextBox? GetLayoutPlaceholderTextBox(int masterIndex, int layoutIndex, PlaceholderValues placeholderType,
            uint? index = null) {
            ThrowIfDisposed();
            SlideLayoutPart layoutPart = GetSlideLayoutPart(masterIndex, layoutIndex);
            Shape? shape = FindLayoutPlaceholderShape(layoutPart, placeholderType, index);
            return shape == null ? null : new PowerPointTextBox(shape);
        }

        /// <summary>
        ///     Ensures a layout placeholder textbox exists, creating it if missing.
        /// </summary>
        public PowerPointTextBox EnsureLayoutPlaceholderTextBox(int masterIndex, int layoutIndex, PlaceholderValues placeholderType,
            uint? index = null, PowerPointLayoutBox? bounds = null, string? name = null) {
            ThrowIfDisposed();
            SlideLayoutPart layoutPart = GetSlideLayoutPart(masterIndex, layoutIndex);
            Shape? shape = FindLayoutPlaceholderShape(layoutPart, placeholderType, index);
            if (shape != null) {
                return new PowerPointTextBox(shape);
            }

            ShapeTree tree = EnsureLayoutShapeTree(layoutPart);
            uint shapeId = GetNextShapeId(tree);
            uint resolvedIndex = index ?? 0U;
            string resolvedName = name ?? $"{placeholderType} Placeholder";
            PowerPointLayoutBox resolvedBounds = bounds ?? GetFallbackPlaceholderBounds(placeholderType);

            Shape created = CreateLayoutPlaceholderShape(shapeId, resolvedName, placeholderType, resolvedIndex, resolvedBounds);
            tree.AppendChild(created);
            return new PowerPointTextBox(created);
        }

        /// <summary>
        ///     Sets layout placeholder bounds.
        /// </summary>
        public void SetLayoutPlaceholderBounds(int masterIndex, int layoutIndex, PlaceholderValues placeholderType,
            PowerPointLayoutBox bounds, uint? index = null, bool createIfMissing = false) {
            ThrowIfDisposed();
            PowerPointTextBox? textBox = createIfMissing
                ? EnsureLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index, bounds)
                : GetLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index);

            if (textBox == null) {
                throw new InvalidOperationException("Layout placeholder was not found.");
            }

            textBox.Bounds = bounds;
        }

        /// <summary>
        ///     Sets layout placeholder text margins in centimeters.
        /// </summary>
        public void SetLayoutPlaceholderTextMarginsCm(int masterIndex, int layoutIndex, PlaceholderValues placeholderType,
            double leftCm, double topCm, double rightCm, double bottomCm, uint? index = null, bool createIfMissing = false) {
            ThrowIfDisposed();
            PowerPointTextBox? textBox = createIfMissing
                ? EnsureLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index)
                : GetLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index);

            if (textBox == null) {
                throw new InvalidOperationException("Layout placeholder was not found.");
            }

            textBox.SetTextMarginsCm(leftCm, topCm, rightCm, bottomCm);
        }

        /// <summary>
        ///     Sets layout placeholder text margins in inches.
        /// </summary>
        public void SetLayoutPlaceholderTextMarginsInches(int masterIndex, int layoutIndex, PlaceholderValues placeholderType,
            double leftInches, double topInches, double rightInches, double bottomInches, uint? index = null,
            bool createIfMissing = false) {
            ThrowIfDisposed();
            PowerPointTextBox? textBox = createIfMissing
                ? EnsureLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index)
                : GetLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index);

            if (textBox == null) {
                throw new InvalidOperationException("Layout placeholder was not found.");
            }

            textBox.SetTextMarginsInches(leftInches, topInches, rightInches, bottomInches);
        }

        /// <summary>
        ///     Sets layout placeholder text margins in points.
        /// </summary>
        public void SetLayoutPlaceholderTextMarginsPoints(int masterIndex, int layoutIndex, PlaceholderValues placeholderType,
            double leftPoints, double topPoints, double rightPoints, double bottomPoints, uint? index = null,
            bool createIfMissing = false) {
            ThrowIfDisposed();
            PowerPointTextBox? textBox = createIfMissing
                ? EnsureLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index)
                : GetLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index);

            if (textBox == null) {
                throw new InvalidOperationException("Layout placeholder was not found.");
            }

            textBox.SetTextMarginsPoints(leftPoints, topPoints, rightPoints, bottomPoints);
        }

        /// <summary>
        ///     Sets layout placeholder text styling and optional bullet settings.
        /// </summary>
        public void SetLayoutPlaceholderTextStyle(int masterIndex, int layoutIndex, PlaceholderValues placeholderType,
            PowerPointTextStyle style, uint? index = null, int? level = null, char? bulletChar = null,
            A.TextAutoNumberSchemeValues? numbering = null, bool createIfMissing = false) {
            ThrowIfDisposed();
            PowerPointTextBox? textBox = createIfMissing
                ? EnsureLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index)
                : GetLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index);

            if (textBox == null) {
                throw new InvalidOperationException("Layout placeholder was not found.");
            }

            PowerPointParagraph paragraph = textBox.Paragraphs.FirstOrDefault() ?? textBox.AddParagraph();
            if (level != null) {
                paragraph.Level = level;
            }

            if (numbering != null) {
                paragraph.SetNumbered(numbering.Value);
            } else if (bulletChar != null) {
                paragraph.SetBullet(bulletChar.Value);
            }

            style.Apply(paragraph);
        }

        /// <summary>
        ///     Ensures a native footer placeholder exists on the specified slide layout.
        /// </summary>
        public PowerPointTextBox EnsureLayoutFooterPlaceholderTextBox(int masterIndex = 0, int layoutIndex = 0,
            string? text = null, PowerPointLayoutBox? bounds = null, uint? index = null) {
            return EnsureLayoutHeaderFooterPlaceholderTextBox(masterIndex, layoutIndex, PlaceholderValues.Footer,
                text, bounds, index ?? 10U, "Footer Placeholder");
        }

        /// <summary>
        ///     Ensures a native date/time placeholder exists on the specified slide layout.
        /// </summary>
        public PowerPointTextBox EnsureLayoutDateTimePlaceholderTextBox(int masterIndex = 0, int layoutIndex = 0,
            string? text = null, PowerPointLayoutBox? bounds = null, uint? index = null) {
            return EnsureLayoutHeaderFooterPlaceholderTextBox(masterIndex, layoutIndex, PlaceholderValues.DateAndTime,
                text, bounds, index ?? 11U, "Date Placeholder");
        }

        /// <summary>
        ///     Ensures a native slide-number placeholder exists on the specified slide layout.
        /// </summary>
        public PowerPointTextBox EnsureLayoutSlideNumberPlaceholderTextBox(int masterIndex = 0, int layoutIndex = 0,
            string? text = null, PowerPointLayoutBox? bounds = null, uint? index = null) {
            return EnsureLayoutHeaderFooterPlaceholderTextBox(masterIndex, layoutIndex, PlaceholderValues.SlideNumber,
                text, bounds, index ?? 12U, "Slide Number Placeholder");
        }

        /// <summary>
        ///     Ensures native footer, date/time, and slide-number placeholders exist on the specified slide layout.
        /// </summary>
        public IReadOnlyList<PowerPointTextBox> EnsureLayoutHeaderFooterPlaceholders(int masterIndex = 0, int layoutIndex = 0,
            string? footerText = null, string? dateTimeText = null, string? slideNumberText = null) {
            return new[] {
                EnsureLayoutFooterPlaceholderTextBox(masterIndex, layoutIndex, footerText),
                EnsureLayoutDateTimePlaceholderTextBox(masterIndex, layoutIndex, dateTimeText),
                EnsureLayoutSlideNumberPlaceholderTextBox(masterIndex, layoutIndex, slideNumberText)
            };
        }

        private static Shape? FindLayoutPlaceholderShape(SlideLayoutPart layoutPart, PlaceholderValues placeholderType, uint? index) {
            ShapeTree? shapeTree = layoutPart.SlideLayout?.CommonSlideData?.ShapeTree;
            if (shapeTree == null) {
                return null;
            }

            foreach (OpenXmlElement element in shapeTree.ChildElements) {
                PlaceholderShape? placeholder = GetLayoutPlaceholderShape(element);
                if (placeholder?.Type?.Value != placeholderType) {
                    continue;
                }
                if (index != null && placeholder.Index?.Value != index.Value) {
                    continue;
                }
                if (element is Shape shape) {
                    return shape;
                }
            }

            return null;
        }

        private static PlaceholderShape? GetLayoutPlaceholderShape(OpenXmlElement element) {
            return element switch {
                Shape s => s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                DocumentFormat.OpenXml.Presentation.Picture p =>
                    p.NonVisualPictureProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                GraphicFrame g => g.NonVisualGraphicFrameProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape,
                _ => null
            };
        }

        private static ShapeTree EnsureLayoutShapeTree(SlideLayoutPart layoutPart) {
            SlideLayout layout = layoutPart.SlideLayout ??= new SlideLayout();
            CommonSlideData common = layout.CommonSlideData ??= new CommonSlideData(new ShapeTree());
            ShapeTree tree = common.ShapeTree ??= new ShapeTree();

            if (tree.GetFirstChild<NonVisualGroupShapeProperties>() == null) {
                tree.PrependChild(new NonVisualGroupShapeProperties(
                    new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                    new NonVisualGroupShapeDrawingProperties(),
                    new ApplicationNonVisualDrawingProperties()));
            }

            if (tree.GetFirstChild<GroupShapeProperties>() == null) {
                tree.AppendChild(PowerPointUtils.CreateDefaultGroupShapeProperties());
            }

            return tree;
        }

        private static uint GetNextShapeId(ShapeTree shapeTree) {
            uint maxId = shapeTree
                .Descendants<NonVisualDrawingProperties>()
                .Select(properties => properties.Id?.Value ?? 0U)
                .DefaultIfEmpty(0U)
                .Max();
            return maxId + 1U;
        }

        private static Shape CreateLayoutPlaceholderShape(uint id, string name, PlaceholderValues type, uint index,
            PowerPointLayoutBox bounds) {
            return new Shape(
                new NonVisualShapeProperties(
                    new NonVisualDrawingProperties { Id = id, Name = name },
                    new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                    new ApplicationNonVisualDrawingProperties(new PlaceholderShape { Type = type, Index = index })),
                new ShapeProperties(
                    new A.Transform2D(new A.Offset { X = bounds.Left, Y = bounds.Top },
                        new A.Extents { Cx = bounds.Width, Cy = bounds.Height }),
                    new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle }
                ),
                new TextBody(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(string.Empty)))))
            { };
        }

        private static PowerPointLayoutBox GetFallbackPlaceholderBounds(PlaceholderValues placeholderType) {
            if (placeholderType == PlaceholderValues.Title || placeholderType == PlaceholderValues.CenteredTitle) {
                return new PowerPointLayoutBox(838200L, 365125L, 7772400L, 1470025L);
            }
            if (placeholderType == PlaceholderValues.SubTitle) {
                return new PowerPointLayoutBox(838200L, 2174875L, 7772400L, 1470025L);
            }
            return new PowerPointLayoutBox(838200L, 2174875L, 7772400L, 3962400L);
        }

        private PowerPointTextBox EnsureLayoutHeaderFooterPlaceholderTextBox(int masterIndex, int layoutIndex,
            PlaceholderValues placeholderType, string? text, PowerPointLayoutBox? bounds, uint index, string name) {
            ThrowIfDisposed();

            SlideLayoutPart layoutPart = GetSlideLayoutPart(masterIndex, layoutIndex);
            SetHeaderFooterFlag(layoutPart, placeholderType, true);

            PowerPointTextBox textBox = EnsureLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType,
                index, bounds ?? GetDefaultHeaderFooterBounds(placeholderType), name);
            if (text != null) {
                textBox.Text = text;
            }

            return textBox;
        }

        private PowerPointLayoutBox GetDefaultHeaderFooterBounds(PlaceholderValues placeholderType) {
            long slideWidth = SlideSize.WidthEmus;
            long slideHeight = SlideSize.HeightEmus;
            long margin = PowerPointUnits.FromCentimeters(0.6);
            long footerTop = slideHeight - PowerPointUnits.FromCentimeters(0.8);
            long footerHeight = PowerPointUnits.FromCentimeters(0.45);
            long sideWidth = Math.Max(PowerPointUnits.FromCentimeters(2.0), slideWidth / 5);
            long centerWidth = Math.Max(PowerPointUnits.FromCentimeters(4.0), slideWidth / 3);

            if (placeholderType == PlaceholderValues.DateAndTime) {
                return new PowerPointLayoutBox(margin, footerTop, sideWidth, footerHeight);
            }

            if (placeholderType == PlaceholderValues.SlideNumber) {
                return new PowerPointLayoutBox(slideWidth - margin - sideWidth, footerTop, sideWidth, footerHeight);
            }

            return new PowerPointLayoutBox((slideWidth - centerWidth) / 2, footerTop, centerWidth, footerHeight);
        }

        private static void SetHeaderFooterFlag(SlideLayoutPart layoutPart, PlaceholderValues placeholderType, bool visible) {
            HeaderFooter headerFooter = EnsureHeaderFooter(layoutPart);
            if (placeholderType == PlaceholderValues.Footer) {
                headerFooter.Footer = visible;
                return;
            }

            if (placeholderType == PlaceholderValues.DateAndTime) {
                headerFooter.DateTime = visible;
                return;
            }

            if (placeholderType == PlaceholderValues.SlideNumber) {
                headerFooter.SlideNumber = visible;
            }
        }

        private static HeaderFooter EnsureHeaderFooter(SlideLayoutPart layoutPart) {
            SlideLayout layout = layoutPart.SlideLayout ??= new SlideLayout();
            HeaderFooter? headerFooter = layout.GetFirstChild<HeaderFooter>();
            if (headerFooter != null) {
                return headerFooter;
            }

            headerFooter = new HeaderFooter();
            SlideLayoutExtensionList? extensionList = layout.GetFirstChild<SlideLayoutExtensionList>();
            if (extensionList != null) {
                layout.InsertBefore(headerFooter, extensionList);
            } else {
                layout.AppendChild(headerFooter);
            }

            return headerFooter;
        }

    }
}
