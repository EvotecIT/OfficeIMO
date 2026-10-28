using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a paragraph within a textbox.
    /// </summary>
    public partial class PowerPointParagraph {
        private readonly SlidePart? _slidePart;
        private readonly OpenXmlPartContainer? _ownerPart;

        internal PowerPointParagraph(A.Paragraph paragraph, SlidePart? slidePart = null, OpenXmlPartContainer? ownerPart = null) {
            Paragraph = paragraph;
            _slidePart = slidePart;
            _ownerPart = ownerPart ?? slidePart;
        }

        internal A.Paragraph Paragraph { get; }
        internal OpenXmlPartContainer? OwnerPart => _ownerPart;

        /// <summary>
        /// Text content of the paragraph.
        /// </summary>
        public string Text {
            get => string.Concat(InlineNodes.Select(node => node.Text));
            set {
                ClearInlineContent();
                InsertRun(value ?? string.Empty);
            }
        }

        internal void ClearInlineContent() {
            string[] discardedSoundIds = PowerPointEmbeddedSound.GetRelationshipIds(Paragraph);
            Paragraph.RemoveAllChildren<A.Run>();
            Paragraph.RemoveAllChildren<A.Break>();
            Paragraph.RemoveAllChildren<A.Field>();
            PowerPointEmbeddedSound.RemoveIfUnused(
                _ownerPart as OpenXmlPart ?? _slidePart,
                discardedSoundIds);
        }

        /// <summary>
        /// Adds a run to the paragraph and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph AddText(string text, Action<PowerPointTextRun>? configure = null) {
            A.Run run = InsertRun(text);
            var wrapper = new PowerPointTextRun(run, _slidePart, _ownerPart);
            configure?.Invoke(wrapper);
            return this;
        }

        /// <summary>
        /// Adds formatted text to the paragraph and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph AddFormattedText(string text, bool bold = false, bool italic = false,
            PowerPointUnderlineStyle? underline = null, bool strikethrough = false) {
            A.Run run = InsertRun(text);
            var wrapper = new PowerPointTextRun(run, _slidePart, _ownerPart);
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
            return this;
        }

        /// <summary>
        /// Replaces text within the paragraph runs while preserving formatting.
        /// </summary>
        public int ReplaceText(string oldValue, string newValue) {
            if (oldValue == null) {
                throw new ArgumentNullException(nameof(oldValue));
            }
            if (oldValue.Length == 0) {
                throw new ArgumentException("Old value cannot be empty.", nameof(oldValue));
            }

            string replacement = newValue ?? string.Empty;
            int count = 0;

            foreach (A.Run run in Paragraph.Elements<A.Run>()) {
                foreach (A.Text text in run.Elements<A.Text>()) {
                    string current = text.Text ?? string.Empty;
                    int occurrences = CountOccurrences(current, oldValue);
                    if (occurrences == 0) {
                        continue;
                    }

                    text.Text = current.Replace(oldValue, replacement);
                    count += occurrences;
                }
            }

            return count;
        }

        /// <summary>
        /// Runs within the paragraph.
        /// </summary>
        public IReadOnlyList<PowerPointTextRun> Runs =>
            Paragraph.Elements<A.Run>().Select(r => new PowerPointTextRun(r, _slidePart, _ownerPart)).ToList();

        /// <summary>
        /// Adds a run to the paragraph.
        /// </summary>
        public PowerPointTextRun AddRun(string text, Action<PowerPointTextRun>? configure = null) {
            A.Run run = InsertRun(text);
            var wrapper = new PowerPointTextRun(run, _slidePart, _ownerPart);
            configure?.Invoke(wrapper);
            return wrapper;
        }

        /// <summary>
        /// Gets or sets paragraph alignment.
        /// </summary>
        public PowerPointTextAlignment? Alignment {
            get => Paragraph.ParagraphProperties?.Alignment?.Value.ToOfficeEnum();
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.Alignment = value?.ToOpenXml();
            }
        }

        /// <summary>
        /// Gets or sets the explicit right-to-left paragraph direction.
        /// </summary>
        public bool? RightToLeft {
            get => Paragraph.ParagraphProperties?.RightToLeft?.Value;
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.RightToLeft = value;
            }
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
            get => Paragraph.ParagraphProperties?.Level?.Value;
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.Level = value;
            }
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
        public string? BulletCharacter => Paragraph.ParagraphProperties?.GetFirstChild<A.CharacterBullet>()?.Char?.Value;

        /// <summary>
        /// Gets whether the paragraph uses PowerPoint auto-numbering.
        /// </summary>
        public bool IsNumbered => Paragraph.ParagraphProperties?.GetFirstChild<A.AutoNumberedBullet>() != null;

        /// <summary>
        /// Gets the PowerPoint automatic-numbering scheme assigned to this paragraph, when present.
        /// </summary>
        public PowerPointNumberingScheme? NumberingScheme =>
            Paragraph.ParagraphProperties?.GetFirstChild<A.AutoNumberedBullet>()?.Type?.Value.ToOfficeEnum();

        /// <summary>
        /// Gets the explicit PowerPoint numbering start value, when present.
        /// </summary>
        public int? NumberingStartAt => Paragraph.ParagraphProperties?.GetFirstChild<A.AutoNumberedBullet>()?.StartAt?.Value;

        /// <summary>
        /// Gets or sets paragraph indentation as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? Indent {
            get => FromTextCoordinate(Paragraph.ParagraphProperties?.Indent?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.Indent = ToTextCoordinate(value);
            }
        }

        /// <summary>
        /// Gets or sets paragraph left margin as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LeftMargin {
            get => FromTextCoordinate(Paragraph.ParagraphProperties?.LeftMargin?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.LeftMargin = ToTextCoordinate(value);
            }
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
            A.ParagraphProperties props = EnsureParagraphProperties();
            props.LeftMargin = margin;
            props.Indent = indent;
            return this;
        }

        /// <summary>
        /// Gets or sets line spacing as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? LineSpacing {
            get => FromSpacingPoints(Paragraph.ParagraphProperties?.LineSpacing?.SpacingPoints?.Val?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                if (value == null) {
                    props.LineSpacing = null;
                    return;
                }
                props.LineSpacing = new A.LineSpacing(new A.SpacingPoints { Val = ToSpacingPoints(value.Value) });
            }
        }

        /// <summary>
        /// Gets or sets line spacing as a multiplier (1.0 = 100%).
        /// </summary>
        public double? LineSpacingMultiplier {
            get => FromSpacingPercent(Paragraph.ParagraphProperties?.LineSpacing?.SpacingPercent?.Val?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                if (value == null) {
                    props.LineSpacing = null;
                    return;
                }
                props.LineSpacing = new A.LineSpacing(new A.SpacingPercent { Val = ToSpacingPercent(value.Value) });
            }
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
            get => FromSpacingPoints(Paragraph.ParagraphProperties?.SpaceBefore?.SpacingPoints?.Val?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                if (value == null) {
                    props.SpaceBefore = null;
                    return;
                }
                props.SpaceBefore = new A.SpaceBefore(new A.SpacingPoints { Val = ToSpacingPoints(value.Value) });
            }
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
            get => FromSpacingPoints(Paragraph.ParagraphProperties?.SpaceAfter?.SpacingPoints?.Val?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                if (value == null) {
                    props.SpaceAfter = null;
                    return;
                }
                props.SpaceAfter = new A.SpaceAfter(new A.SpacingPoints { Val = ToSpacingPoints(value.Value) });
            }
        }

        /// <summary>
        /// Sets space after the paragraph and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetSpaceAfter(global::OfficeIMO.OfficeLength value) {
            SpaceAfter = value;
            return this;
        }

        /// <summary>
        /// Sets the current run bold property and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetBold(bool isBold = true) {
            PowerPointTextRun run = GetDefaultRun();
            run.Bold = isBold;
            return this;
        }

        /// <summary>
        /// Sets the current run italic property and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetItalic(bool isItalic = true) {
            PowerPointTextRun run = GetDefaultRun();
            run.Italic = isItalic;
            return this;
        }

        /// <summary>
        /// Sets the current run underline property and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetUnderline(bool underline = true) {
            PowerPointTextRun run = GetDefaultRun();
            run.Underline = underline;
            return this;
        }

        /// <summary>
        /// Sets the current run strikethrough property and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetStrikethrough(bool strikethrough = true) {
            PowerPointTextRun run = GetDefaultRun();
            run.Strikethrough = strikethrough;
            return this;
        }

        /// <summary>
        /// Sets the current run font size and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetFontSize(global::OfficeIMO.OfficeLength size) {
            PowerPointTextRun run = GetDefaultRun();
            run.FontSize = size;
            return this;
        }

        /// <summary>
        /// Sets the current run font name and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetFontName(string fontName) {
            PowerPointTextRun run = GetDefaultRun();
            run.FontName = fontName;
            return this;
        }

        /// <summary>
        /// Sets the current run color and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetColor(string color) {
            PowerPointTextRun run = GetDefaultRun();
            run.Color = color;
            return this;
        }

        /// <summary>
        /// Sets the current run highlight color and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetHighlightColor(string color) {
            PowerPointTextRun run = GetDefaultRun();
            run.HighlightColor = color;
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
        public string? BulletFontName {
            get => Paragraph.ParagraphProperties?.GetFirstChild<A.BulletFont>()?.Typeface;
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.RemoveAllChildren<A.BulletFont>();
                if (!string.IsNullOrWhiteSpace(value)) {
                    InsertParagraphPropertyChild(props, new A.BulletFont { Typeface = value });
                }
            }
        }

        /// <summary>
        /// Gets or sets the bullet size as a physical length.
        /// </summary>
        public global::OfficeIMO.OfficeLength? BulletSize {
            get {
                int? size = Paragraph.ParagraphProperties?.GetFirstChild<A.BulletSizePoints>()?.Val?.Value;
                return size.HasValue ? global::OfficeIMO.OfficeLength.Points(size.Value / 100D) : (global::OfficeIMO.OfficeLength?)null;
            }
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.RemoveAllChildren<A.BulletSizePoints>();
                props.RemoveAllChildren<A.BulletSizePercentage>();
                if (value != null) {
                    InsertParagraphPropertyChild(props, new A.BulletSizePoints { Val = ToSpacingPoints(value.Value) });
                }
            }
        }

        /// <summary>
        /// Gets or sets the bullet size in percent (100 = 100%).
        /// </summary>
        public int? BulletSizePercent {
            get {
                int? size = Paragraph.ParagraphProperties?.GetFirstChild<A.BulletSizePercentage>()?.Val?.Value;
                return size != null ? size / 1000 : null;
            }
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.RemoveAllChildren<A.BulletSizePercentage>();
                props.RemoveAllChildren<A.BulletSizePoints>();
                if (value != null) {
                    InsertParagraphPropertyChild(props, new A.BulletSizePercentage { Val = value.Value * 1000 });
                }
            }
        }

        /// <summary>
        /// Sets the bullet font and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetBulletFont(string fontName) {
            BulletFontName = fontName;
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
        public void SetBullet(char bulletChar = '\u2022') {
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            InsertParagraphPropertyChild(props, new A.CharacterBullet { Char = bulletChar.ToString() });
        }

        /// <summary>
        /// Applies an auto-numbered bullet to the paragraph.
        /// </summary>
        public void SetNumbered(PowerPointNumberingScheme style, int startAt = 1) {
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            InsertParagraphPropertyChild(props, new A.AutoNumberedBullet { Type = style.ToOpenXml(), StartAt = startAt });
        }

        /// <summary>
        /// Applies an auto-numbered bullet to the paragraph without resetting the start value.
        /// </summary>
        public void SetNumbered(PowerPointNumberingScheme style) {
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            InsertParagraphPropertyChild(props, new A.AutoNumberedBullet { Type = style.ToOpenXml() });
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
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            InsertParagraphPropertyChild(props, new A.NoBullet());
        }

        private PowerPointTextRun GetDefaultRun() {
            A.Run run = Paragraph.Elements<A.Run>().LastOrDefault() ?? InsertRun(string.Empty);
            return new PowerPointTextRun(run, _slidePart);
        }

        private A.Run InsertRun(string text) {
            A.Run run = new(new A.Text(text ?? string.Empty));
            A.EndParagraphRunProperties? endProps = Paragraph.GetFirstChild<A.EndParagraphRunProperties>();
            if (endProps != null) {
                Paragraph.InsertBefore(run, endProps);
            } else {
                Paragraph.Append(run);
            }
            return run;
        }

        private static void ClearBulletInternal(A.ParagraphProperties props) {
            props.RemoveAllChildren<A.BulletFont>();
            props.RemoveAllChildren<A.BulletSizePoints>();
            props.RemoveAllChildren<A.BulletSizePercentage>();
            props.RemoveAllChildren<A.CharacterBullet>();
            props.RemoveAllChildren<A.AutoNumberedBullet>();
            props.RemoveAllChildren<A.NoBullet>();
        }

        private static void InsertParagraphPropertyChild(A.ParagraphProperties props, OpenXmlElement child) {
            int childOrder = GetParagraphPropertyChildOrder(child);
            OpenXmlElement? insertBefore = props.ChildElements
                .FirstOrDefault(existing => GetParagraphPropertyChildOrder(existing) > childOrder);

            if (insertBefore == null) {
                props.Append(child);
            } else {
                props.InsertBefore(child, insertBefore);
            }
        }

        private static int GetParagraphPropertyChildOrder(OpenXmlElement element) {
            return element switch {
                A.LineSpacing => 10,
                A.SpaceBefore => 20,
                A.SpaceAfter => 30,
                A.BulletColorText => 40,
                A.BulletColor => 50,
                A.BulletSizeText => 60,
                A.BulletSizePercentage => 70,
                A.BulletSizePoints => 80,
                A.BulletFontText => 90,
                A.BulletFont => 100,
                A.NoBullet => 110,
                A.AutoNumberedBullet => 120,
                A.CharacterBullet => 130,
                A.TabStopList => 140,
                A.DefaultRunProperties => 150,
                A.ExtensionList => 160,
                _ => 1000
            };
        }

        private A.ParagraphProperties EnsureParagraphProperties() {
            return Paragraph.ParagraphProperties ??= new A.ParagraphProperties();
        }

        private static int CountOccurrences(string value, string oldValue) {
            int count = 0;
            int index = 0;
            while (true) {
                index = value.IndexOf(oldValue, index, StringComparison.Ordinal);
                if (index < 0) {
                    break;
                }
                count++;
                index += oldValue.Length;
            }
            return count;
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
