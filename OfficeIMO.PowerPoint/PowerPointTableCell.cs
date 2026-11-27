using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a cell within a PowerPoint table.
    /// </summary>
    public partial class PowerPointTableCell {
        private const int D = OfficeXmlNamespaces.DrawingML;
        private const int EmusPerPoint = 12700;

        private readonly PowerPointTable _table;
        private readonly PowerPointHeldNode _cell;

        // A cell (a:tc) of the table, held where the table holds it.
        internal PowerPointTableCell(PowerPointTable table, OfficeXmlNode cell) {
            _table = table;
            _cell = new PowerPointHeldNode(cell, table);
        }

        /// <summary>The cell's element as it is held now.</summary>
        internal OfficeXmlNode Element => _cell.Element;

        /// <summary>The table the cell belongs to.</summary>
        internal PowerPointTable Table => _table;

        /// <summary>The part that holds the cell.</summary>
        internal PowerPointPartRef OwnerPartRef => _table.OwnerPartRef;

        /// <summary>The slide the cell is on, or none for a table of a layout or master.</summary>
        internal PowerPointPartRef SlidePartRef => _table.SlidePartRef;

        private void MarkEdited() => OwnerPartRef.MarkEdited();

        private OfficeTheme? ThemeFor(OfficeColor? color) =>
            PowerPointDrawingColor.NeedsTheme(color) ? PowerPointDrawingColor.ThemeFor(OwnerPartRef, _table.OwnerSlide) : null;

        private OfficeXmlNode BodyElement => Element.Child(D, "txBody");

        private OfficeXmlNode PropertiesElement => Element.Child(D, "tcPr");

        /// <summary>
        ///     Gets or sets the text contained in the cell.
        /// </summary>
        public string Text {
            get => PowerPointParagraph.ReadBodyText(EnsureTextBody());
            set {
                OfficeXmlNode body = EnsureTextBody();
                // The first paragraph's properties, its first run's formatting and its end properties carry over, without
                // the hyperlinks the old text had.
                OfficeXmlNode source = body.Child(D, "p");
                OfficeXmlNode paragraphProperties = source.Child(D, "pPr");
                if (!paragraphProperties.IsNone) {
                    paragraphProperties = paragraphProperties.Clone();
                    foreach (OfficeXmlNode link in paragraphProperties.Descendants(D, "hlinkClick")) link.Remove();
                    foreach (OfficeXmlNode link in paragraphProperties.Descendants(D, "hlinkMouseOver")) link.Remove();
                }
                OfficeXmlNode endProperties = source.Child(D, "endParaRPr");
                endProperties = endProperties.IsNone ? body.CreateElement(D, "endParaRPr") : endProperties.Clone();
                WithoutHyperlinks(endProperties);
                string language = endProperties.Attribute("lang") ?? PowerPointTableTextDefaults.Language;
                endProperties.SetAttribute("lang", language);
                OfficeXmlNode runProperties = source.Child(D, "r").Child(D, "rPr");
                runProperties = runProperties.IsNone ? body.CreateElement(D, "rPr") : runProperties.Clone();
                WithoutHyperlinks(runProperties);
                if (runProperties.Attribute("lang") == null) runProperties.SetAttribute("lang", language);

                OfficeXmlNode paragraph = body.CreateElement(D, "p");
                if (!paragraphProperties.IsNone) paragraph.Append(paragraphProperties);
                PowerPointTextContent.InsertInlineContent(paragraph, value, runProperties, default);
                paragraph.Append(endProperties);

                PowerPointPartRef owner = OwnerPartRef, slide = SlidePartRef;
                foreach (OfficeXmlNode previous in body.Children(D, "p").ToList()) PowerPointTextRun.DiscardContent(previous, owner, slide);
                body.Append(paragraph);
                MarkEdited();
            }
        }

        private static void WithoutHyperlinks(OfficeXmlNode properties) {
            properties.RemoveChildren(D, "hlinkClick");
            properties.RemoveChildren(D, "hlinkMouseOver");
        }

        /// <summary>
        ///     Gets or sets the merge information for this cell.
        ///     Tuple is in format (rows, columns).
        /// </summary>
        public (int rows, int columns) Merge {
            get {
                OfficeXmlNode cell = Element;
                return (cell.Int32Attribute("rowSpan") ?? 1, cell.Int32Attribute("gridSpan") ?? 1);
            }
            set {
                OfficeXmlNode cell = Element;
                cell.SetInt32Attribute("rowSpan", value.rows <= 1 ? null : value.rows);
                cell.SetInt32Attribute("gridSpan", value.columns <= 1 ? null : value.columns);
                MarkEdited();
            }
        }

        /// <summary>
        ///     Gets or sets the horizontal alignment of the cell text.
        /// </summary>
        public PowerPointTextAlignment? HorizontalAlignment {
            get => BodyElement.Child(D, "p").Child(D, "pPr").Attribute("algn") is { } token
                && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointTextAlignment alignment) ? alignment : null;
            set {
                string? token = value?.ToOpenXmlToken();
                OfficeXmlNode body = BodyElement;
                if (body.IsNone) {
                    body = CreateBody(withParagraph: false);
                    body.Append(body.CreateElement(D, "p"));
                }
                OfficeXmlNode paragraph = body.Child(D, "p");
                if (paragraph.IsNone) paragraph = body.Append(body.CreateElement(D, "p"));
                paragraph.GetOrPrependChild(D, "pPr").SetAttribute("algn", token);
                MarkEdited();
            }
        }

        /// <summary>
        ///     Gets or sets the local border color (all sides), retaining native theme slots and opacity.
        /// </summary>
        public OfficeColor? BorderColor {
            get => PowerPointDrawingColor.Read(PropertiesElement.Child(D, "lnL").Child(D, "solidFill"));
            set {
                if (value == null) {
                    ClearBorders(PowerPointTableCellBorders.All);
                    return;
                }

                SetBorders(PowerPointTableCellBorders.All, value.Value);
            }
        }

        /// <summary>
        ///     Gets a value indicating whether this cell is a merged continuation cell.
        /// </summary>
        public bool IsMergedCell {
            get {
                OfficeXmlNode cell = Element;
                return cell.BooleanAttribute("hMerge") == true || cell.BooleanAttribute("vMerge") == true;
            }
        }

        /// <summary>
        ///     Gets a value indicating whether this cell is the anchor for a merged range.
        /// </summary>
        public bool IsMergeAnchor {
            get {
                OfficeXmlNode cell = Element;
                return (cell.Int32Attribute("rowSpan") ?? 1) > 1 || (cell.Int32Attribute("gridSpan") ?? 1) > 1;
            }
        }

        /// <summary>
        ///     Replaces literal text in the cell paragraphs.
        /// </summary>
        /// <param name="text">Text to find.</param>
        /// <param name="replacement">Replacement text; <c>null</c> removes each match.</param>
        /// <param name="options">Comparison settings. <c>null</c> matches ordinally, ignoring case, inside or across words.</param>
        /// <returns>The number of replaced matches.</returns>
        /// <remarks>A match can span adjacent runs within one paragraph. Replacement text takes the formatting of the run where the match starts.</remarks>
        public int ReplaceAll(string text, string replacement, OfficeTextSearchOptions? options = null) {
            OfficeTextSearch.ValidateSearchText(text, nameof(text));
            return BodyElement.IsNone ? 0 : Paragraphs.Sum(paragraph => paragraph.ReplaceAll(text, replacement, options));
        }

        /// <summary>
        ///     Gets formatted runs from the first cell paragraph.
        /// </summary>
        public IReadOnlyList<PowerPointTextRun> Runs {
            get {
                var runs = new List<PowerPointTextRun>();
                PowerPointPartRef slide = SlidePartRef, owner = OwnerPartRef;
                foreach (OfficeXmlNode run in BodyElement.Child(D, "p").Children(D, "r")) runs.Add(new PowerPointTextRun(run, slide, owner, owningSlide: _table.OwnerSlide));
                return runs;
            }
        }

        /// <summary>
        ///     Appends text as a new run in the first cell paragraph and returns that run.
        /// </summary>
        public PowerPointTextRun AddText(string text) => ParagraphFor(EnsureParagraph()).InsertRun(text);

        /// <summary>
        ///     Gets or sets whether the cell text is bold. Setting applies to every run and line break in the cell.
        /// </summary>
        public bool Bold {
            get => FirstRunProperties.BooleanAttribute("b") == true;
            set => FormatRuns(properties => properties.SetAttribute("b", value ? "1" : "0"));
        }

        /// <summary>
        ///     Gets or sets whether the cell text is italic.
        /// </summary>
        public bool Italic {
            get => FirstRunProperties.BooleanAttribute("i") == true;
            set => FormatRuns(properties => properties.SetAttribute("i", value ? "1" : "0"));
        }

        /// <summary>
        ///     Gets or sets the font size as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? FontSize {
            get {
                int? size = FirstRunProperties.Int32Attribute("sz");
                return size.HasValue ? global::OfficeIMO.OfficeLength.Points(size.Value / 100D) : (global::OfficeIMO.OfficeLength?)null;
            }
            set {
                int? drawingSize = PowerPointTextDefaults.ToDrawingFontSize(value?.ToPoints(), nameof(value));
                FormatRuns(properties => properties.SetInt32Attribute("sz", drawingSize));
            }
        }

        /// <summary>
        ///     Gets or sets the font family.
        /// </summary>
        public string? FontFamily {
            get => FirstRunProperties.Child(D, "latin").Attribute("typeface");
            set => FormatRuns(properties => PowerPointTextPropertyFormatting.SetFontFamily(properties, value));
        }

        /// <summary>
        ///     Gets or sets the text color. Alpha below 255 is written as native DrawingML opacity.
        /// </summary>
        public global::OfficeIMO.OfficeColor? Color {
            get => PowerPointDrawingColor.Read(FirstRunProperties.Child(D, "solidFill"));
            set {
                OfficeTheme? theme = ThemeFor(value);
                FormatRuns(properties => PowerPointTextPropertyFormatting.SetColor(properties, value, theme));
            }
        }

        /// <summary>
        ///     Gets or sets the left padding as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? PaddingLeft {
            get => GetMargin("marL");
            set => SetMargin("marL", value);
        }

        /// <summary>
        ///     Gets or sets the right padding as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? PaddingRight {
            get => GetMargin("marR");
            set => SetMargin("marR", value);
        }

        /// <summary>
        ///     Gets or sets the top padding as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? PaddingTop {
            get => GetMargin("marT");
            set => SetMargin("marT", value);
        }

        /// <summary>
        ///     Gets or sets the bottom padding as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? PaddingBottom {
            get => GetMargin("marB");
            set => SetMargin("marB", value);
        }

        private global::OfficeIMO.OfficeLength? GetMargin(string attribute) =>
            PropertiesElement.Int32Attribute(attribute) is int value ? global::OfficeIMO.OfficeLength.Emus(value) : (global::OfficeIMO.OfficeLength?)null;

        private void SetMargin(string attribute, global::OfficeIMO.OfficeLength? value) {
            EnsureProperties().SetInt32Attribute(attribute, ToEmus(value?.ToPoints(), EmusPerPoint));
            MarkEdited();
        }

        /// <summary>
        ///     Gets or sets the vertical alignment of the cell text (top/center/bottom).
        /// </summary>
        public PowerPointTextVerticalAlignment? VerticalAlignment {
            get => PropertiesElement.Attribute("anchor") is { } token
                && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointTextVerticalAlignment alignment) ? alignment : null;
            set {
                string? token = value?.ToOpenXmlToken();
                EnsureProperties().SetAttribute("anchor", token);
                MarkEdited();
            }
        }

        /// <summary>
        ///     Gets or sets the text auto-fit behavior for the cell.
        /// </summary>
        public PowerPointTextAutoFit? TextAutoFit {
            get {
                OfficeXmlNode body = BodyElement.Child(D, "bodyPr");
                if (body.IsNone) return null;
                if (!body.Child(D, "noAutofit").IsNone) return PowerPointTextAutoFit.None;
                if (!body.Child(D, "normAutofit").IsNone) return PowerPointTextAutoFit.Normal;
                if (!body.Child(D, "spAutoFit").IsNone) return PowerPointTextAutoFit.Shape;
                return null;
            }
            set {
                OfficeXmlNode body = EnsureBodyProperties();
                RemoveAutoFit(body);
                if (value != null) {
                    string name = value.Value switch {
                        PowerPointTextAutoFit.None => "noAutofit",
                        PowerPointTextAutoFit.Normal => "normAutofit",
                        _ => "spAutoFit"
                    };
                    body.AddOrdered(D, name, ATextBodyPropertiesView.Order);
                }
                MarkEdited();
            }
        }

        private static void RemoveAutoFit(OfficeXmlNode body) {
            body.RemoveChildren(D, "noAutofit");
            body.RemoveChildren(D, "normAutofit");
            body.RemoveChildren(D, "spAutoFit");
        }

        /// <summary>
        ///     Gets or sets detailed auto-fit options (only applies to Normal auto-fit).
        /// </summary>
        public PowerPointTextAutoFitOptions? TextAutoFitOptions {
            get {
                OfficeXmlNode normal = BodyElement.Child(D, "bodyPr").Child(D, "normAutofit");
                if (normal.IsNone) return null;
                return PowerPointTextAutoFitOptions.FromOpenXmlValues(normal.Int32Attribute("fontScale"), normal.Int32Attribute("lnSpcReduction"));
            }
            set {
                if (value == null) {
                    BodyElement.Child(D, "bodyPr").RemoveChildren(D, "normAutofit");
                    MarkEdited();
                    return;
                }

                OfficeXmlNode body = EnsureBodyProperties();
                body.RemoveChildren(D, "noAutofit");
                body.RemoveChildren(D, "spAutoFit");
                ApplyNormalAutoFitOptions(body.GetOrAddChild(D, "normAutofit", ATextBodyPropertiesView.Order), value.Value);
                MarkEdited();
            }
        }

        /// <summary>
        ///     Sets the auto-fit mode and optional Normal auto-fit options.
        /// </summary>
        public PowerPointTableCell SetTextAutoFit(PowerPointTextAutoFit fit, PowerPointTextAutoFitOptions? options = null) {
            TextAutoFit = fit;
            if (fit == PowerPointTextAutoFit.Normal && options != null) {
                ApplyNormalAutoFitOptions(EnsureBodyProperties().GetOrAddChild(D, "normAutofit", ATextBodyPropertiesView.Order), options.Value);
                MarkEdited();
            }
            return this;
        }

        // The cell's border elements in tcPr, by side.
        private static readonly (PowerPointTableCellBorders Side, string Name)[] BorderElements = {
            (PowerPointTableCellBorders.Left, "lnL"), (PowerPointTableCellBorders.Right, "lnR"),
            (PowerPointTableCellBorders.Top, "lnT"), (PowerPointTableCellBorders.Bottom, "lnB"),
            (PowerPointTableCellBorders.DiagonalDown, "lnTlToBr"), (PowerPointTableCellBorders.DiagonalUp, "lnBlToTr")
        };

        /// <summary>
        ///     Applies border styling, retaining the existing width or dash when omitted.
        /// </summary>
        public void SetBorders(PowerPointTableCellBorders borders, OfficeColor color, OfficeLength? width = null, PowerPointLineDashStyle? dash = null) {
            ValidateBorderStyle(borders, color, width);
            string? dashToken = dash?.ToOpenXmlToken();
            int? lineWidth = width.HasValue ? PowerPointShape.ToDrawingLineWidth(width.Value, nameof(width)) : null;
            OfficeTheme? theme = ThemeFor(color);

            OfficeXmlNode properties = EnsureProperties();
            foreach ((PowerPointTableCellBorders side, string name) in BorderElements) {
                if (!borders.HasFlag(side)) continue;
                OfficeXmlNode line = properties.GetOrAddChild(D, name, ATableCellPropertiesView.Order);
                SetBorderLineFill(line, PowerPointDrawingColor.CreateFill(line, color, theme));
                if (lineWidth != null) line.SetInt32Attribute("w", lineWidth);
                if (dashToken != null) {
                    line.RemoveChildren(D, "prstDash");
                    line.AddOrdered(D, "prstDash", ALinePropertiesView.Order).SetAttribute("val", dashToken);
                }
            }
            MarkEdited();
        }

        /// <summary>
        ///     Clears borders on the specified sides.
        /// </summary>
        public void ClearBorders(PowerPointTableCellBorders borders) {
            OfficeXmlNode properties = EnsureProperties();
            foreach ((PowerPointTableCellBorders side, string name) in BorderElements) {
                if (borders.HasFlag(side)) properties.RemoveChildren(D, name);
            }
            MarkEdited();
        }

        private static int? ToEmus(double? value, double emusPerUnit) {
            return value != null ? (int)Math.Round(value.Value * emusPerUnit) : null;
        }

        private static void ApplyNormalAutoFitOptions(OfficeXmlNode normal, PowerPointTextAutoFitOptions options) {
            normal.SetInt32Attribute("fontScale", options.FontScaleValue);
            normal.SetInt32Attribute("lnSpcReduction", options.LineSpaceReductionValue);
        }

        private OfficeXmlNode EnsureProperties() => Element.GetOrAddChild(D, "tcPr", ATableCellView.Order);

        // A text body before the cell properties: body and list-style properties, with one empty paragraph when asked.
        private OfficeXmlNode CreateBody(bool withParagraph) {
            OfficeXmlNode cell = Element;
            OfficeXmlNode body = cell.CreateElement(D, "txBody");
            body.Append(body.CreateElement(D, "bodyPr"));
            body.Append(body.CreateElement(D, "lstStyle"));
            if (withParagraph) body.Append(PowerPointTableTextDefaults.CreateParagraph(body));
            body = cell.InsertOrdered(body, ATableCellView.Order);
            MarkEdited();
            return body;
        }

        private OfficeXmlNode EnsureBodyProperties() {
            OfficeXmlNode body = BodyElement;
            if (body.IsNone) return CreateBody(withParagraph: false).Child(D, "bodyPr");
            return body.GetOrPrependChild(D, "bodyPr");
        }

        private OfficeXmlNode FirstRunProperties => BodyElement.Child(D, "p").Child(D, "r").Child(D, "rPr");

        private OfficeXmlNode EnsureParagraph() {
            OfficeXmlNode body = BodyElement;
            if (body.IsNone) body = CreateBody(withParagraph: false);
            OfficeXmlNode paragraph = body.Child(D, "p");
            if (paragraph.IsNone) {
                paragraph = body.Append(body.CreateElement(D, "p"));
                MarkEdited();
            }
            return paragraph;
        }

        // A paragraph wrapper held the way the cell is.
        private PowerPointParagraph ParagraphFor(OfficeXmlNode paragraph) => new(paragraph, SlidePartRef, OwnerPartRef, _table.OwnerSlide);

        private void FormatRuns(Action<OfficeXmlNode> change) {
            if (BodyElement.Child(D, "p").Child(D, "r").IsNone) ParagraphFor(EnsureParagraph()).InsertRun(string.Empty);
            PowerPointTextContent.FormatRuns(BodyElement, change);
            MarkEdited();
        }
    }
}
