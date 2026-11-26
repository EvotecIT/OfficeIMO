using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.PowerPoint.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a paragraph within a textbox.
    /// </summary>
    public partial class PowerPointParagraph : IPackageSdkBinding {
        private OfficeXmlNode _element;
        private readonly PowerPointPartRef _initialSlidePart;
        private readonly PowerPointPartRef _initialOwnerPart;
        private readonly PowerPointSlide? _packageSlide;
        private readonly PowerPointPartRef _packagePart;

        /// <summary>
        /// A paragraph (<c>a:p</c>) of a part, wherever the part is held. <paramref name="slidePart"/> is the slide the
        /// text belongs to and <paramref name="ownerPart"/> the part that holds it.
        /// </summary>
        internal PowerPointParagraph(OfficeXmlNode paragraph, PowerPointPartRef slidePart, PowerPointPartRef ownerPart) {
            _element = paragraph;
            _initialSlidePart = !slidePart.IsNone ? slidePart : ownerPart.Kind == PowerPointPartKind.Slide ? ownerPart : default;
            _initialOwnerPart = !ownerPart.IsNone ? ownerPart : slidePart;
            if (paragraph.Tree != null && !_initialOwnerPart.IsNone && !_initialOwnerPart.IsSdk && paragraph.IsAttached) {
                _packagePart = _initialOwnerPart;
                _initialOwnerPart.Document?.RegisterPackageBinding(this);
            }
        }

        // A paragraph of a slide's package tree; it binds to the same paragraph in the SDK when the SDK opens.
        internal PowerPointParagraph(PowerPointSlide slide, OfficeXmlTree tree, int node) {
            _packageSlide = slide;
            _packagePart = _initialSlidePart = _initialOwnerPart = slide.PartRef;
            _element = new OfficeXmlNode(tree, node);
            slide.OwnerDocument?.RegisterPackageBinding(this);
        }

        /// <summary>The paragraph's element as it is held now: once the SDK opened, a package paragraph is its SDK element.</summary>
        internal OfficeXmlNode Element {
            get {
                // A removed slide's part name can be reused by a new slide, so its content never binds.
                if (_element.Tree != null && !_packagePart.IsNone && _packagePart.Document is { IsSdkLoaded: true } && _packageSlide is not { RemovedFromPackage: true }) {
                    _element = OfficeXmlNode.FromSdk(PowerPointSlide.BindPackageNode(_packagePart, _element.Tree, _element.Node));
                }
                return _element;
            }
        }

        /// <summary>The paragraph's package tree while it works on one, or null.</summary>
        internal OfficeXmlTree? PackageTree => Element.Tree;

        /// <summary>The paragraph's node in <see cref="PackageTree"/>.</summary>
        internal int PackageNode => Element.Node;

        void IPackageSdkBinding.BindSdk() {
            try {
                _ = Element;
            } catch (InvalidOperationException) {
                // A paragraph the opened slide doesn't have where expected reports it when it is used.
            }
        }

        /// <summary>The slide the paragraph is on, found from where the element is now when it can be.</summary>
        internal PowerPointPartRef SlidePartRef {
            get {
                PowerPointPartRef current = ContainingPart(Element);
                return current.Kind == PowerPointPartKind.Slide ? current : _initialSlidePart;
            }
        }

        /// <summary>The part that holds the paragraph, found from where the element is now when it can be.</summary>
        internal PowerPointPartRef OwnerPartRef {
            get {
                PowerPointPartRef current = ContainingPart(Element);
                return current.Kind == PowerPointPartKind.Slide ? current : !_initialOwnerPart.IsNone ? _initialOwnerPart : _initialSlidePart;
            }
        }

        private static PowerPointPartRef ContainingPart(OfficeXmlNode element) =>
            element.IsSdk ? PowerPointPartRef.FromSdk(OfficeSdkXml.PartOf(element.SdkElement!)) : default;

        private void MarkEdited() => OwnerPartRef.MarkEdited();

        // A run wrapper for a child of this paragraph, bound the same way the paragraph is.
        // A node taken from the package tree binds by its place there, even when the SDK opened after it was taken.
        private PowerPointTextRun RunFor(OfficeXmlNode element) =>
            _packageSlide != null && element.Tree != null
                ? new PowerPointTextRun(_packageSlide, element.Tree, element.Node)
                : new PowerPointTextRun(element, SlidePartRef, OwnerPartRef);

        /// <summary>
        /// Text content of the paragraph.
        /// </summary>
        public string Text {
            get => ReadText(Element);
            set {
                ClearInlineContent();
                InsertRun(value ?? string.Empty);
            }
        }

        /// <summary>Reads the same text as <see cref="InlineNodes" /> without creating run wrappers.</summary>
        internal static string ReadText(OfficeXmlNode paragraph) {
            string? single = null;
            System.Text.StringBuilder? builder = null;
            foreach (OfficeXmlNode child in paragraph.Children()) {
                if (child.NamespaceId != OfficeXmlNamespaces.DrawingML) continue;
                string? part = child.LocalName switch {
                    "r" or "fld" => child.ChildText(OfficeXmlNamespaces.DrawingML, "t") ?? string.Empty,
                    "br" => Environment.NewLine,
                    _ => null
                };
                if (part == null) continue;
                if (builder != null) {
                    builder.Append(part);
                } else if (single == null) {
                    single = part;
                } else {
                    builder = new System.Text.StringBuilder(single).Append(part);
                }
            }
            return builder?.ToString() ?? single ?? string.Empty;
        }

        /// <summary>Joins paragraph text with line breaks, as the text box and table cell <c>Text</c> getters return it.</summary>
        internal static string ReadBodyText(OfficeXmlNode textBody) {
            string? single = null;
            System.Text.StringBuilder? builder = null;
            foreach (OfficeXmlNode paragraph in textBody.Children(OfficeXmlNamespaces.DrawingML, "p")) {
                string text = ReadText(paragraph);
                if (builder != null) {
                    builder.Append(Environment.NewLine).Append(text);
                } else if (single == null) {
                    single = text;
                } else {
                    builder = new System.Text.StringBuilder(single).Append(Environment.NewLine).Append(text);
                }
            }
            return builder?.ToString() ?? single ?? string.Empty;
        }

        internal void ClearInlineContent() {
            OfficeXmlNode paragraph = Element;
            var soundIds = new List<string>();
            foreach (OfficeXmlNode child in paragraph.Children().ToList()) {
                if (child.NamespaceId != OfficeXmlNamespaces.DrawingML || child.LocalName is not ("r" or "br" or "fld")) continue;
                CollectSoundIds(child, soundIds);
                child.Remove();
            }
            MarkEdited();
            PowerPointPartRef owner = OwnerPartRef;
            if (owner.IsNone) return;
            foreach (string id in soundIds) owner.RemoveRelationshipIfUnused(id, default);
        }

        private static void CollectSoundIds(OfficeXmlNode element, List<string> ids) {
            foreach (OfficeXmlNode child in element.Children()) {
                if (child.Is(OfficeXmlNamespaces.DrawingML, "snd") && child.Attribute(OfficeXmlNamespaces.Relationships, "embed") is { Length: > 0 } id && !ids.Contains(id)) ids.Add(id);
                CollectSoundIds(child, ids);
            }
        }

        /// <summary>
        /// Appends text as a new run and returns that run. Formatting the returned run formats only that text;
        /// line breaks in <paramref name="text"/> stay part of the same run.
        /// </summary>
        public PowerPointTextRun AddText(string text) => InsertRun(text);

        /// <summary>
        /// Appends text as a new run with basic formatting and returns that run.
        /// </summary>
        public PowerPointTextRun AddFormattedText(string text, bool bold = false, bool italic = false,
            PowerPointUnderlineStyle? underline = null, bool strikethrough = false) {
            PowerPointTextRun wrapper = InsertRun(text);
            if (bold) {
                wrapper.Bold = true;
            }
            if (italic) {
                wrapper.Italic = true;
            }
            if (underline != null) {
                wrapper.UnderlineStyle = underline;
            }
            if (strikethrough) {
                wrapper.Strikethrough = true;
            }
            return wrapper;
        }

        /// <summary>
        /// Runs within the paragraph.
        /// </summary>
        public IReadOnlyList<PowerPointTextRun> Runs {
            get {
                var runs = new List<PowerPointTextRun>();
                foreach (OfficeXmlNode run in Element.Children(OfficeXmlNamespaces.DrawingML, "r")) runs.Add(RunFor(run));
                return runs;
            }
        }

        private OfficeXmlNode Properties => Element.Child(OfficeXmlNamespaces.DrawingML, "pPr");

        private OfficeXmlNode EnsureProperties() {
            OfficeXmlNode properties = Element.GetOrAddChild(OfficeXmlNamespaces.DrawingML, "pPr", ATextParagraphView.Order);
            MarkEdited();
            return properties;
        }

        private void SetProperty(string attribute, string? value) => EnsureProperties().SetAttribute(attribute, value);

        /// <summary>
        /// Gets or sets paragraph alignment.
        /// </summary>
        public PowerPointTextAlignment? Alignment {
            get => Properties.Attribute("algn") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointTextAlignment alignment)
                ? alignment : null;
            set => SetProperty("algn", value?.ToOpenXmlToken());
        }

        /// <summary>
        /// Gets or sets the explicit right-to-left paragraph direction.
        /// </summary>
        public bool? RightToLeft {
            get => Properties.BooleanAttribute("rtl");
            set => SetProperty("rtl", OfficeXmlViewValues.Format(value));
        }

        /// <summary>
        /// Sets paragraph alignment and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetAlignment(PowerPointTextAlignment alignment) {
            Alignment = alignment;
            return this;
        }

        /// <summary>
        /// Gets or sets the bullet/list level (0-8).
        /// </summary>
        public int? Level {
            get => Properties.Int32Attribute("lvl");
            set => SetProperty("lvl", OfficeXmlViewValues.Format(value));
        }

        /// <summary>
        /// Sets the bullet/list level and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetLevel(int level) {
            Level = level;
            return this;
        }

        /// <summary>
        /// Gets the character bullet assigned to this paragraph, when present.
        /// </summary>
        public string? BulletCharacter => Properties.Child(OfficeXmlNamespaces.DrawingML, "buChar").Attribute("char");

        private OfficeXmlNode AutoNumber => Properties.Child(OfficeXmlNamespaces.DrawingML, "buAutoNum");

        /// <summary>
        /// Gets whether the paragraph uses PowerPoint auto-numbering.
        /// </summary>
        public bool IsNumbered => !AutoNumber.IsNone;

        /// <summary>
        /// Gets the PowerPoint automatic-numbering scheme assigned to this paragraph, when present.
        /// </summary>
        public PowerPointNumberingScheme? NumberingScheme =>
            AutoNumber.Attribute("type") is { } token && PowerPointOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out PowerPointNumberingScheme scheme)
                ? scheme : null;

        /// <summary>
        /// Gets the explicit PowerPoint numbering start value, when present.
        /// </summary>
        public int? NumberingStartAt => AutoNumber.Int32Attribute("startAt");

        /// <summary>
        /// Gets or sets paragraph indentation as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? Indent {
            get => FromTextCoordinate(Properties.Int32Attribute("indent"));
            set => SetProperty("indent", OfficeXmlViewValues.Format(ToTextCoordinate(value)));
        }

        /// <summary>
        /// Gets or sets paragraph left margin as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LeftMargin {
            get => FromTextCoordinate(Properties.Int32Attribute("marL"));
            set => SetProperty("marL", OfficeXmlViewValues.Format(ToTextCoordinate(value)));
        }

        /// <summary>
        /// Sets indentation and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetIndent(global::OfficeIMO.OfficeLength value) {
            Indent = value;
            return this;
        }

        /// <summary>
        /// Sets left margin and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetLeftMargin(global::OfficeIMO.OfficeLength value) {
            LeftMargin = value;
            return this;
        }

        /// <summary>
        /// Sets a hanging indent and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetHanging(global::OfficeIMO.OfficeLength value) {
            int margin = checked((int)value.ToEmus());
            int indent = checked(-margin);
            OfficeXmlNode properties = EnsureProperties();
            properties.SetInt32Attribute("marL", margin);
            properties.SetInt32Attribute("indent", indent);
            return this;
        }

        // The a:spcPts or a:spcPct value inside a spacing element (a:lnSpc, a:spcBef, a:spcAft).
        private int? SpacingValue(string spacing, string kind) =>
            Properties.Child(OfficeXmlNamespaces.DrawingML, spacing).Child(OfficeXmlNamespaces.DrawingML, kind).Int32Attribute("val");

        // Replaces a spacing element with one holding a single a:spcPts or a:spcPct; null removes it.
        private void SetSpacing(string spacing, string kind, int? value) {
            OfficeXmlNode properties = EnsureProperties();
            properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, spacing);
            if (value == null) return;
            OfficeXmlNode element = properties.CreateElement(OfficeXmlNamespaces.DrawingML, spacing);
            element.Append(element.CreateElement(OfficeXmlNamespaces.DrawingML, kind)).SetInt32Attribute("val", value);
            properties.InsertOrdered(element, ATextParagraphPropertiesView.Order);
        }

        /// <summary>
        /// Gets or sets line spacing as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LineSpacing {
            get => FromSpacingPoints(SpacingValue("lnSpc", "spcPts"));
            set => SetSpacing("lnSpc", "spcPts", value == null ? null : ToSpacingPoints(value.Value));
        }

        /// <summary>
        /// Gets or sets line spacing as a multiplier (1.0 = 100%).
        /// </summary>
        public double? LineSpacingMultiplier {
            get => FromSpacingPercent(SpacingValue("lnSpc", "spcPct"));
            set => SetSpacing("lnSpc", "spcPct", value == null ? null : ToSpacingPercent(value.Value));
        }

        /// <summary>
        /// Sets line spacing and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetLineSpacing(global::OfficeIMO.OfficeLength value) {
            LineSpacing = value;
            return this;
        }

        /// <summary>
        /// Sets line spacing as a multiplier (1.0 = 100%) and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetLineSpacingMultiplier(double multiplier) {
            LineSpacingMultiplier = multiplier;
            return this;
        }

        /// <summary>
        /// Gets or sets space before the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? SpaceBefore {
            get => FromSpacingPoints(SpacingValue("spcBef", "spcPts"));
            set => SetSpacing("spcBef", "spcPts", value == null ? null : ToSpacingPoints(value.Value));
        }

        /// <summary>
        /// Sets space before the paragraph and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetSpaceBefore(global::OfficeIMO.OfficeLength value) {
            SpaceBefore = value;
            return this;
        }

        /// <summary>
        /// Gets or sets space after the paragraph as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? SpaceAfter {
            get => FromSpacingPoints(SpacingValue("spcAft", "spcPts"));
            set => SetSpacing("spcAft", "spcPts", value == null ? null : ToSpacingPoints(value.Value));
        }

        /// <summary>
        /// Sets space after the paragraph and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetSpaceAfter(global::OfficeIMO.OfficeLength value) {
            SpaceAfter = value;
            return this;
        }

        /// <summary>
        /// Sets the bold property of every run in the paragraph and returns the paragraph for chaining.
        /// Use the run returned by <see cref="AddText(string)"/> to format a single run.
        /// </summary>
        public PowerPointParagraph SetBold(bool isBold = true) {
            FormatRuns(run => run.Bold = isBold);
            return this;
        }

        /// <summary>
        /// Sets the italic property of every run in the paragraph and returns the paragraph for chaining.
        /// Use the run returned by <see cref="AddText(string)"/> to format a single run.
        /// </summary>
        public PowerPointParagraph SetItalic(bool isItalic = true) {
            FormatRuns(run => run.Italic = isItalic);
            return this;
        }

        /// <summary>
        /// Sets the underline property of every run in the paragraph and returns the paragraph for chaining.
        /// Use the run returned by <see cref="AddText(string)"/> to format a single run.
        /// </summary>
        public PowerPointParagraph SetUnderline(bool underline = true) {
            FormatRuns(run => run.Underline = underline);
            return this;
        }

        /// <summary>
        /// Sets the strikethrough property of every run in the paragraph and returns the paragraph for chaining.
        /// Use the run returned by <see cref="AddText(string)"/> to format a single run.
        /// </summary>
        public PowerPointParagraph SetStrikethrough(bool strikethrough = true) {
            FormatRuns(run => run.Strikethrough = strikethrough);
            return this;
        }

        /// <summary>
        /// Sets the font size of every run in the paragraph and returns the paragraph for chaining.
        /// Use the run returned by <see cref="AddText(string)"/> to format a single run.
        /// </summary>
        public PowerPointParagraph SetFontSize(global::OfficeIMO.OfficeLength size) {
            FormatRuns(run => run.FontSize = size);
            return this;
        }

        /// <summary>
        /// Sets the font family of every run in the paragraph and returns the paragraph for chaining.
        /// Use the run returned by <see cref="AddText(string)"/> to format a single run.
        /// </summary>
        public PowerPointParagraph SetFontFamily(string fontFamily) {
            FormatRuns(run => run.FontFamily = fontFamily);
            return this;
        }

        /// <summary>
        /// Sets the color of every run in the paragraph and returns the paragraph for chaining.
        /// Use the run returned by <see cref="AddText(string)"/> to format a single run.
        /// </summary>
        public PowerPointParagraph SetColor(global::OfficeIMO.OfficeColor color) {
            FormatRuns(run => run.Color = color);
            return this;
        }

        /// <summary>
        /// Sets the background (highlight) color of every run in the paragraph and returns the paragraph for chaining.
        /// Use the run returned by <see cref="AddText(string)"/> to format a single run.
        /// </summary>
        public PowerPointParagraph SetBackgroundColor(global::OfficeIMO.OfficeColor color) {
            FormatRuns(run => run.BackgroundColor = color);
            return this;
        }

        /// <summary>
        /// Sets the paragraph text and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetText(string text) {
            Text = text;
            return this;
        }

        /// <summary>
        /// Gets or sets the bullet font name.
        /// </summary>
        public string? BulletFontFamily {
            get => Properties.Child(OfficeXmlNamespaces.DrawingML, "buFont").Attribute("typeface");
            set {
                OfficeXmlNode properties = EnsureProperties();
                properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "buFont");
                if (!string.IsNullOrWhiteSpace(value)) {
                    properties.AddOrdered(OfficeXmlNamespaces.DrawingML, "buFont", ATextParagraphPropertiesView.Order).SetAttribute("typeface", value);
                }
            }
        }

        /// <summary>
        /// Gets or sets the bullet size as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? BulletSize {
            get {
                int? size = Properties.Child(OfficeXmlNamespaces.DrawingML, "buSzPts").Int32Attribute("val");
                return size.HasValue ? global::OfficeIMO.OfficeLength.Points(size.Value / 100D) : (global::OfficeIMO.OfficeLength?)null;
            }
            set {
                OfficeXmlNode properties = EnsureProperties();
                properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "buSzPts");
                properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "buSzPct");
                if (value != null) {
                    properties.AddOrdered(OfficeXmlNamespaces.DrawingML, "buSzPts", ATextParagraphPropertiesView.Order).SetInt32Attribute("val", ToSpacingPoints(value.Value));
                }
            }
        }

        /// <summary>
        /// Gets or sets the bullet size in percent (100 = 100%).
        /// </summary>
        public int? BulletSizePercent {
            get {
                int? size = Properties.Child(OfficeXmlNamespaces.DrawingML, "buSzPct").Int32Attribute("val");
                return size != null ? size / 1000 : null;
            }
            set {
                OfficeXmlNode properties = EnsureProperties();
                properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "buSzPct");
                properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, "buSzPts");
                if (value != null) {
                    properties.AddOrdered(OfficeXmlNamespaces.DrawingML, "buSzPct", ATextParagraphPropertiesView.Order).SetInt32Attribute("val", value.Value * 1000);
                }
            }
        }

        /// <summary>
        /// Sets the bullet font and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetBulletFontFamily(string fontFamily) {
            BulletFontFamily = fontFamily;
            return this;
        }

        /// <summary>
        /// Sets the bullet size and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetBulletSize(global::OfficeIMO.OfficeLength size) {
            BulletSize = size;
            return this;
        }

        /// <summary>
        /// Sets the bullet size in percent (100 = 100%) and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetBulletSizePercent(int percent) {
            BulletSizePercent = percent;
            return this;
        }

        /// <summary>
        /// Applies a character bullet to the paragraph.
        /// </summary>
        public void SetBullet(char bulletChar = '•') {
            OfficeXmlNode properties = EnsureProperties();
            ClearBulletInternal(properties);
            properties.AddOrdered(OfficeXmlNamespaces.DrawingML, "buChar", ATextParagraphPropertiesView.Order).SetAttribute("char", bulletChar.ToString());
        }

        /// <summary>
        /// Applies an auto-numbered bullet to the paragraph.
        /// </summary>
        public void SetNumbered(PowerPointNumberingScheme style, int startAt = 1) {
            OfficeXmlNode numbered = SetAutoNumber(style);
            numbered.SetInt32Attribute("startAt", startAt);
        }

        /// <summary>
        /// Applies an auto-numbered bullet to the paragraph without resetting the start value.
        /// </summary>
        public void SetNumbered(PowerPointNumberingScheme style) => SetAutoNumber(style);

        private OfficeXmlNode SetAutoNumber(PowerPointNumberingScheme style) {
            OfficeXmlNode properties = EnsureProperties();
            ClearBulletInternal(properties);
            OfficeXmlNode numbered = properties.AddOrdered(OfficeXmlNamespaces.DrawingML, "buAutoNum", ATextParagraphPropertiesView.Order);
            numbered.SetAttribute("type", style.ToOpenXmlToken());
            return numbered;
        }

        /// <summary>
        /// Applies a default auto-numbered bullet (Arabic period) to the paragraph.
        /// </summary>
        public void SetNumbered(int startAt = 1) {
            SetNumbered(PowerPointNumberingScheme.ArabicPeriod, startAt);
        }

        /// <summary>
        /// Clears any bullet/numbering from the paragraph.
        /// </summary>
        public void ClearBullet() {
            OfficeXmlNode properties = EnsureProperties();
            ClearBulletInternal(properties);
            properties.AddOrdered(OfficeXmlNamespaces.DrawingML, "buNone", ATextParagraphPropertiesView.Order);
        }

        /// <summary>
        /// Applies <paramref name="format"/> to every run in the paragraph. A paragraph without runs gets one empty,
        /// formatted run, as PowerPoint stores an empty formatted line.
        /// </summary>
        internal void FormatRuns(Action<PowerPointTextRun> format) {
            List<OfficeXmlNode> runs = Element.Children(OfficeXmlNamespaces.DrawingML, "r").ToList();
            if (runs.Count == 0) {
                format(InsertRun(string.Empty));
                return;
            }
            foreach (OfficeXmlNode run in runs) {
                format(RunFor(run));
            }
        }

        /// <summary>
        /// Adds a run with <paramref name="text"/> before the end paragraph properties, carrying a copy of
        /// <paramref name="properties"/> when given, and returns it.
        /// </summary>
        internal PowerPointTextRun InsertRun(string? text, OfficeXmlNode properties = default) {
            OfficeXmlNode paragraph = Element;
            OfficeXmlNode run = paragraph.CreateElement(OfficeXmlNamespaces.DrawingML, "r");
            run.Append(run.CreateElement(OfficeXmlNamespaces.DrawingML, "t"));
            run = paragraph.InsertBefore(run, paragraph.Child(OfficeXmlNamespaces.DrawingML, "endParaRPr"));
            PowerPointTextContent.ApplyProperties(run, properties);
            MarkEdited();
            PowerPointTextRun wrapper = RunFor(run);
            wrapper.Text = text ?? string.Empty;
            return wrapper;
        }

        private static readonly string[] BulletElements = { "buFont", "buSzPts", "buSzPct", "buChar", "buAutoNum", "buNone" };

        private static void ClearBulletInternal(OfficeXmlNode properties) {
            foreach (string name in BulletElements) properties.RemoveChildren(OfficeXmlNamespaces.DrawingML, name);
        }

        private static int? ToTextCoordinate(global::OfficeIMO.OfficeLength? length) {
            return length != null
                ? checked((int)length.Value.ToEmus())
                : null;
        }

        private static global::OfficeIMO.OfficeLength? FromTextCoordinate(int? value) {
            return value.HasValue ? global::OfficeIMO.OfficeLength.Emus(value.Value) : (global::OfficeIMO.OfficeLength?)null;
        }

        private static int ToSpacingPoints(global::OfficeIMO.OfficeLength length) {
            return (int)Math.Round(length.ToPoints() * 100);
        }

        private static global::OfficeIMO.OfficeLength? FromSpacingPoints(int? value) {
            return value.HasValue ? global::OfficeIMO.OfficeLength.Points(value.Value / 100D) : (global::OfficeIMO.OfficeLength?)null;
        }

        private static int ToSpacingPercent(double multiplier) {
            return (int)Math.Round(multiplier * 100000d);
        }

        private static double? FromSpacingPercent(int? value) {
            return value != null ? value.Value / 100000d : null;
        }

    }
}
