using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const int LayoutPresentationNamespace = OfficeXmlNamespaces.PresentationML;

        // A master's layouts in relationship order, as the index-based layout APIs count them.
        private List<PowerPointPartRef> LayoutPartRefs(int masterIndex) =>
            RequireMasterPartRef(masterIndex).RelatedParts(PowerPointPartRef.SlideLayoutRelationshipType).ToList();

        // The layout at an index of a master, throwing as an out-of-range index does.
        private PowerPointPartRef RequireLayoutPartRef(int masterIndex, int layoutIndex) {
            List<PowerPointPartRef> layouts = LayoutPartRefs(masterIndex);
            ArgumentOutOfRangeException.ThrowIfNegative(layoutIndex);
            ArgumentOutOfRangeException.ThrowIfGreaterThanOrEqual(layoutIndex, layouts.Count);
            return layouts[layoutIndex];
        }

        // The master at an index, throwing as an out-of-range index does.
        private PowerPointPartRef RequireMasterPartRef(int masterIndex) {
            ArgumentOutOfRangeException.ThrowIfNegative(masterIndex);
            PowerPointPartRef master = MasterPartRef(masterIndex);
            if (master.IsNone) throw new ArgumentOutOfRangeException(nameof(masterIndex));
            return master;
        }

        private static string LayoutName(PowerPointPartRef layout) =>
            layout.Root.Child(LayoutPresentationNamespace, "cSld").Attribute("name") ?? string.Empty;

        // A layout's declared type; one the schema doesn't know reads as none.
        private static PowerPointSlideLayoutType? LayoutTypeOf(PowerPointPartRef layout) =>
            PowerPointSlideLayoutTypeTokens.TryParse(layout.Root.Attribute("type"), out PowerPointSlideLayoutType type) ? type : null;

        /// <summary>
        ///     Returns the layouts available for a slide master.
        /// </summary>
        public IReadOnlyList<PowerPointSlideLayoutInfo> GetSlideLayouts(int masterIndex = 0) {
            ThrowIfDisposed();
            PowerPointPartRef master = RequireMasterPartRef(masterIndex);
            List<PowerPointPartRef> layouts = LayoutPartRefs(masterIndex);
            var results = new List<PowerPointSlideLayoutInfo>(layouts.Count);
            for (int i = 0; i < layouts.Count; i++) {
                results.Add(new PowerPointSlideLayoutInfo(masterIndex, i, LayoutName(layouts[i]), LayoutTypeOf(layouts[i]), master.RelationshipIdOf(layouts[i])));
            }

            return results;
        }

        /// <summary>
        ///     Finds a layout index by layout type.
        /// </summary>
        public int GetLayoutIndex(PowerPointSlideLayoutType layoutType, int masterIndex = 0) {
            ThrowIfDisposed();
            int found = LayoutPartRefs(masterIndex).FindIndex(layout => LayoutTypeOf(layout) == layoutType);
            if (found >= 0) return found;
            throw new InvalidOperationException($"Layout type '{layoutType}' not found for master {masterIndex}.");
        }

        /// <summary>
        ///     Finds a layout index by layout name.
        /// </summary>
        public int GetLayoutIndex(string layoutName, int masterIndex = 0, bool ignoreCase = true) {
            ThrowIfDisposed();
            ArgumentNullException.ThrowIfNull(layoutName);
            StringComparison comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
            int found = LayoutPartRefs(masterIndex).FindIndex(layout => string.Equals(LayoutName(layout), layoutName, comparison));
            if (found >= 0) return found;
            throw new InvalidOperationException($"Layout '{layoutName}' not found for master {masterIndex}.");
        }

        /// <summary>
        ///     Adds a slide using a layout type.
        /// </summary>
        public PowerPointSlide AddSlide(PowerPointSlideLayoutType layoutType, int masterIndex = 0) {
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

        /// <summary>
        ///     Retrieves a layout placeholder textbox for a master/layout pair.
        /// </summary>
        public PowerPointTextBox? GetLayoutPlaceholderTextBox(int masterIndex, int layoutIndex, PowerPointPlaceholderType placeholderType,
            uint? index = null) {
            ThrowIfDisposed();
            PowerPointSlideLayout layout = GetLayout(RequireLayoutPartRef(masterIndex, layoutIndex));
            OfficeXmlNode shape = FindLayoutPlaceholderShape(layout, placeholderType, index);
            return shape.IsNone ? null : (PowerPointTextBox)layout.WrapShape(shape);
        }

        /// <summary>Reads layout geometry in native inheritance order before a partial edit materializes a local transform.</summary>
        internal PowerPointLayoutBox? GetEffectiveLayoutPlaceholderBounds(int masterIndex, int layoutIndex,
            PowerPointPlaceholderType placeholderType, uint? index = null) {
            ThrowIfDisposed();
            PowerPointPartRef layoutPart = RequireLayoutPartRef(masterIndex, layoutIndex);
            OfficeXmlNode shape = FindLayoutPlaceholderShape(GetLayout(layoutPart), placeholderType, index);
            OfficeXmlNode placeholder = shape.Child(LayoutPresentationNamespace, "nvSpPr").Child(LayoutPresentationNamespace, "nvPr").Child(LayoutPresentationNamespace, "ph");
            if (shape.IsNone || placeholder.IsNone) return null;
            return LocalBounds(shape) ?? PowerPointSlide.GetExportPlaceholderBounds(layoutPart, placeholder);
        }

        // The offset and extents of a layout shape's own transform, when all four are set.
        private static PowerPointLayoutBox? LocalBounds(OfficeXmlNode shape) {
            OfficeXmlNode transform = shape.Child(LayoutPresentationNamespace, "spPr").Child(OfficeXmlNamespaces.DrawingML, "xfrm");
            OfficeXmlNode offset = transform.Child(OfficeXmlNamespaces.DrawingML, "off"), extents = transform.Child(OfficeXmlNamespaces.DrawingML, "ext");
            long? x = offset.Int64Attribute("x"), y = offset.Int64Attribute("y"), cx = extents.Int64Attribute("cx"), cy = extents.Int64Attribute("cy");
            if (x == null || y == null || cx == null || cy == null) return null;
            return new PowerPointLayoutBox(OfficeLength.Emus(x.Value), OfficeLength.Emus(y.Value), OfficeLength.Emus(cx.Value), OfficeLength.Emus(cy.Value));
        }

        /// <summary>
        ///     Ensures a layout placeholder textbox exists, creating it if missing.
        /// </summary>
        public PowerPointTextBox EnsureLayoutPlaceholderTextBox(int masterIndex, int layoutIndex, PowerPointPlaceholderType placeholderType,
            uint? index = null, PowerPointLayoutBox? bounds = null, string? name = null) {
            ThrowIfDisposed();
            PowerPointSlideLayout layout = GetLayout(RequireLayoutPartRef(masterIndex, layoutIndex));
            OfficeXmlNode shape = FindLayoutPlaceholderShape(layout, placeholderType, index);
            if (!shape.IsNone) {
                return (PowerPointTextBox)layout.WrapShape(shape);
            }

            OfficeXmlNode tree = layout.ShapeTree;
            uint shapeId = GetNextShapeId(tree);
            string typeToken = placeholderType.ToOpenXmlToken();
            uint resolvedIndex = index ?? 0U;
            string resolvedName = name ?? $"{placeholderType} Placeholder";
            PowerPointLayoutBox resolvedBounds = bounds ?? GetFallbackPlaceholderBounds(placeholderType);

            OfficeXmlNode created = tree.Append(CreateLayoutPlaceholderShape(tree, shapeId, resolvedName, typeToken, resolvedIndex, resolvedBounds));
            layout.PartRef.MarkEdited();
            return (PowerPointTextBox)layout.WrapShape(created);
        }

        /// <summary>
        ///     Sets layout placeholder bounds.
        /// </summary>
        public void SetLayoutPlaceholderBounds(int masterIndex, int layoutIndex, PowerPointPlaceholderType placeholderType,
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
        ///     Sets layout placeholder physical text margins.
        /// </summary>
        public void SetLayoutPlaceholderTextMargins(int masterIndex, int layoutIndex, PowerPointPlaceholderType placeholderType,
            global::OfficeIMO.OfficeLength left, global::OfficeIMO.OfficeLength top, global::OfficeIMO.OfficeLength right, global::OfficeIMO.OfficeLength bottom, uint? index = null,
            bool createIfMissing = false) {
            ThrowIfDisposed();
            PowerPointTextBox? textBox = createIfMissing
                ? EnsureLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index)
                : GetLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType, index);

            if (textBox == null) {
                throw new InvalidOperationException("Layout placeholder was not found.");
            }

            textBox.SetTextMargins(left, top, right, bottom);
        }

        /// <summary>
        ///     Sets layout placeholder text styling and optional bullet settings.
        /// </summary>
        public void SetLayoutPlaceholderTextStyle(int masterIndex, int layoutIndex, PowerPointPlaceholderType placeholderType,
            PowerPointTextStyle style, uint? index = null, int? level = null, char? bulletChar = null,
            PowerPointNumberingScheme? numbering = null, bool createIfMissing = false) {
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
            if (style.HasFormatting) {
                OfficeTheme? theme = style.Color.HasValue || style.BackgroundColor.HasValue
                    ? PowerPointDrawingColor.ThemeFor(paragraph.OwnerPartRef, null) : null;
                OfficeXmlNode element = paragraph.Element;
                OfficeXmlNode properties = element.GetOrPrependChild(OfficeXmlNamespaces.DrawingML, "pPr");
                OfficeXmlNode defaults = properties.GetOrAddChild(OfficeXmlNamespaces.DrawingML, "defRPr", ATextParagraphPropertiesView.Order);
                style.ApplyDefaults(defaults, theme);
                paragraph.OwnerPartRef.MarkEdited();
            }
        }

        /// <summary>
        ///     Ensures a native footer placeholder exists on the specified slide layout.
        /// </summary>
        public PowerPointTextBox EnsureLayoutFooterPlaceholderTextBox(int masterIndex = 0, int layoutIndex = 0,
            string? text = null, PowerPointLayoutBox? bounds = null, uint? index = null) {
            return EnsureLayoutHeaderFooterPlaceholderTextBox(masterIndex, layoutIndex, PowerPointPlaceholderType.Footer,
                text, bounds, index ?? 10U, "Footer Placeholder");
        }

        /// <summary>
        ///     Ensures a native date/time placeholder exists on the specified slide layout.
        /// </summary>
        public PowerPointTextBox EnsureLayoutDateTimePlaceholderTextBox(int masterIndex = 0, int layoutIndex = 0,
            string? text = null, PowerPointLayoutBox? bounds = null, uint? index = null) {
            return EnsureLayoutHeaderFooterPlaceholderTextBox(masterIndex, layoutIndex, PowerPointPlaceholderType.DateAndTime,
                text, bounds, index ?? 11U, "Date Placeholder");
        }

        /// <summary>
        ///     Ensures a native slide-number placeholder exists on the specified slide layout.
        /// </summary>
        public PowerPointTextBox EnsureLayoutSlideNumberPlaceholderTextBox(int masterIndex = 0, int layoutIndex = 0,
            string? text = null, PowerPointLayoutBox? bounds = null, uint? index = null) {
            return EnsureLayoutHeaderFooterPlaceholderTextBox(masterIndex, layoutIndex, PowerPointPlaceholderType.SlideNumber,
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

        // The first shape (p:sp) of the layout whose placeholder has the type and, when given, the index.
        private static OfficeXmlNode FindLayoutPlaceholderShape(PowerPointSlideLayout layout, PowerPointPlaceholderType placeholderType, uint? index) {
            OfficeXmlNode tree = layout.Root.Child(LayoutPresentationNamespace, "cSld").Child(LayoutPresentationNamespace, "spTree");
            string typeToken = placeholderType.ToOpenXmlToken();
            foreach (OfficeXmlNode element in tree.Children()) {
                OfficeXmlNode placeholder = LayoutPlaceholderOf(element);
                // A placeholder that declares no type matches none, as the model reads it.
                if (placeholder.Attribute("type") != typeToken) continue;
                if (index != null && placeholder.UInt32Attribute("idx") != index.Value) continue;
                if (element.Is(LayoutPresentationNamespace, "sp")) return element;
            }

            return default;
        }

        private static OfficeXmlNode LayoutPlaceholderOf(OfficeXmlNode element) {
            string? container = element.NamespaceId != LayoutPresentationNamespace ? null : element.LocalName switch {
                "sp" => "nvSpPr",
                "pic" => "nvPicPr",
                "graphicFrame" => "nvGraphicFramePr",
                _ => null
            };
            return container == null ? default
                : element.Child(LayoutPresentationNamespace, container).Child(LayoutPresentationNamespace, "nvPr").Child(LayoutPresentationNamespace, "ph");
        }

        private static uint GetNextShapeId(OfficeXmlNode tree) {
            uint maxId = 0U;
            foreach (OfficeXmlNode drawing in tree.Descendants(LayoutPresentationNamespace, "cNvPr")) maxId = Math.Max(maxId, drawing.UInt32Attribute("id") ?? 0U);
            return maxId + 1U;
        }

        // A placeholder shape as the model adds it to a layout: a rectangle at the bounds with one empty run.
        private static OfficeXmlNode CreateLayoutPlaceholderShape(OfficeXmlNode context, uint id, string name, string typeToken, uint index,
            PowerPointLayoutBox bounds) {
            const int P = LayoutPresentationNamespace, D = OfficeXmlNamespaces.DrawingML;
            OfficeXmlNode shape = context.CreateElement(P, "sp");
            OfficeXmlNode nonVisual = shape.Append(shape.CreateElement(P, "nvSpPr"));
            OfficeXmlNode drawing = nonVisual.Append(nonVisual.CreateElement(P, "cNvPr"));
            drawing.SetUInt32Attribute("id", id);
            drawing.SetAttribute("name", name);
            nonVisual.Append(nonVisual.CreateElement(P, "cNvSpPr")).Append(nonVisual.CreateElement(D, "spLocks")).SetAttribute("noGrp", "1");
            OfficeXmlNode placeholder = nonVisual.Append(nonVisual.CreateElement(P, "nvPr")).Append(nonVisual.CreateElement(P, "ph"));
            placeholder.SetAttribute("type", typeToken);
            placeholder.SetUInt32Attribute("idx", index);
            OfficeXmlNode properties = shape.Append(shape.CreateElement(P, "spPr"));
            OfficeXmlNode transform = properties.Append(properties.CreateElement(D, "xfrm"));
            OfficeXmlNode offset = transform.Append(transform.CreateElement(D, "off"));
            offset.SetInt64Attribute("x", bounds.Left.ToEmus());
            offset.SetInt64Attribute("y", bounds.Top.ToEmus());
            OfficeXmlNode extents = transform.Append(transform.CreateElement(D, "ext"));
            extents.SetInt64Attribute("cx", bounds.Width.ToEmus());
            extents.SetInt64Attribute("cy", bounds.Height.ToEmus());
            OfficeXmlNode geometry = properties.Append(properties.CreateElement(D, "prstGeom"));
            geometry.SetAttribute("prst", "rect");
            geometry.Append(geometry.CreateElement(D, "avLst"));
            OfficeXmlNode body = shape.Append(shape.CreateElement(P, "txBody"));
            body.Append(body.CreateElement(D, "bodyPr"));
            body.Append(body.CreateElement(D, "lstStyle"));
            OfficeXmlNode run = body.Append(body.CreateElement(D, "p")).Append(body.CreateElement(D, "r"));
            run.Append(run.CreateElement(D, "t"));
            return shape;
        }

        private static PowerPointLayoutBox GetFallbackPlaceholderBounds(PowerPointPlaceholderType placeholderType) {
            if (placeholderType == PowerPointPlaceholderType.Title || placeholderType == PowerPointPlaceholderType.CenteredTitle) {
                return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(838200L),OfficeIMO.OfficeLength.Emus(365125L),OfficeIMO.OfficeLength.Emus(7772400L),OfficeIMO.OfficeLength.Emus(1470025L));
            }
            if (placeholderType == PowerPointPlaceholderType.SubTitle) {
                return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(838200L),OfficeIMO.OfficeLength.Emus(2174875L),OfficeIMO.OfficeLength.Emus(7772400L),OfficeIMO.OfficeLength.Emus(1470025L));
            }
            return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(838200L),OfficeIMO.OfficeLength.Emus(2174875L),OfficeIMO.OfficeLength.Emus(7772400L),OfficeIMO.OfficeLength.Emus(3962400L));
        }

        private PowerPointTextBox EnsureLayoutHeaderFooterPlaceholderTextBox(int masterIndex, int layoutIndex,
            PowerPointPlaceholderType placeholderType, string? text, PowerPointLayoutBox? bounds, uint index, string name) {
            ThrowIfDisposed();

            SetHeaderFooterFlag(RequireLayoutPartRef(masterIndex, layoutIndex), placeholderType);

            PowerPointTextBox textBox = EnsureLayoutPlaceholderTextBox(masterIndex, layoutIndex, placeholderType,
                index, bounds ?? GetDefaultHeaderFooterBounds(placeholderType), name);
            if (text != null) {
                textBox.Text = text;
            }

            return textBox;
        }

        private PowerPointLayoutBox GetDefaultHeaderFooterBounds(PowerPointPlaceholderType placeholderType) {
            long slideWidth = SlideSize.Width.ToEmus();
            long slideHeight = SlideSize.Height.ToEmus();
            long margin = PowerPointUnits.FromCentimeters(0.6);
            long footerTop = slideHeight - PowerPointUnits.FromCentimeters(0.8);
            long footerHeight = PowerPointUnits.FromCentimeters(0.45);
            long sideWidth = Math.Max(PowerPointUnits.FromCentimeters(2.0), slideWidth / 5);
            long centerWidth = Math.Max(PowerPointUnits.FromCentimeters(4.0), slideWidth / 3);

            if (placeholderType == PowerPointPlaceholderType.DateAndTime) {
                return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(margin),OfficeIMO.OfficeLength.Emus(footerTop),OfficeIMO.OfficeLength.Emus(sideWidth),OfficeIMO.OfficeLength.Emus(footerHeight));
            }

            if (placeholderType == PowerPointPlaceholderType.SlideNumber) {
                return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus(slideWidth - margin - sideWidth),OfficeIMO.OfficeLength.Emus(footerTop),OfficeIMO.OfficeLength.Emus(sideWidth),OfficeIMO.OfficeLength.Emus(footerHeight));
            }

            return new PowerPointLayoutBox(OfficeIMO.OfficeLength.Emus((slideWidth - centerWidth) / 2),OfficeIMO.OfficeLength.Emus(footerTop),OfficeIMO.OfficeLength.Emus(centerWidth),OfficeIMO.OfficeLength.Emus(footerHeight));
        }

        // Shows the footer, date or slide number on the layout's header/footer settings (p:hf), added before the
        // extension list when missing.
        private static void SetHeaderFooterFlag(PowerPointPartRef layout, PowerPointPlaceholderType placeholderType) {
            string? attribute = placeholderType switch {
                PowerPointPlaceholderType.Footer => "ftr",
                PowerPointPlaceholderType.DateAndTime => "dt",
                PowerPointPlaceholderType.SlideNumber => "sldNum",
                _ => null
            };
            OfficeXmlNode root = layout.Root;
            OfficeXmlNode headerFooter = root.Child(LayoutPresentationNamespace, "hf");
            if (headerFooter.IsNone) {
                headerFooter = root.InsertBefore(root.CreateElement(LayoutPresentationNamespace, "hf"), root.Child(LayoutPresentationNamespace, "extLst"));
            }
            if (attribute != null) headerFooter.SetAttribute(attribute, "1");
            layout.MarkEdited();
        }

    }
}
