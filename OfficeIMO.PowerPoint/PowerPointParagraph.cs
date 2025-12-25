using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Represents a paragraph within a textbox.
    /// </summary>
    public class PowerPointParagraph {
        private readonly SlidePart? _slidePart;

        internal PowerPointParagraph(A.Paragraph paragraph, SlidePart? slidePart = null) {
            Paragraph = paragraph;
            _slidePart = slidePart;
        }

        internal A.Paragraph Paragraph { get; }

        /// <summary>
        /// Text content of the paragraph.
        /// </summary>
        public string Text {
            get => Paragraph.InnerText ?? string.Empty;
            set {
                A.EndParagraphRunProperties? endProps = Paragraph.GetFirstChild<A.EndParagraphRunProperties>();
                endProps?.Remove();
                Paragraph.RemoveAllChildren<A.Run>();
                A.Run run = new(new A.Text(value ?? string.Empty));
                Paragraph.Append(run);
                if (endProps != null) {
                    Paragraph.Append(endProps);
                }
            }
        }

        /// <summary>
        /// Adds a run to the paragraph and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph AddText(string text, Action<PowerPointTextRun>? configure = null) {
            A.Run run = InsertRun(text);
            var wrapper = new PowerPointTextRun(run, _slidePart);
            configure?.Invoke(wrapper);
            return this;
        }

        /// <summary>
        /// Adds formatted text to the paragraph and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph AddFormattedText(string text, bool bold = false, bool italic = false,
            A.TextUnderlineValues? underline = null) {
            A.Run run = InsertRun(text);
            var wrapper = new PowerPointTextRun(run, _slidePart);
            if (bold) {
                wrapper.Bold = true;
            }
            if (italic) {
                wrapper.Italic = true;
            }
            if (underline != null) {
                wrapper.Underline = true;
                wrapper.Run.RunProperties ??= new A.RunProperties();
                wrapper.Run.RunProperties.Underline = underline.Value;
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
            Paragraph.Elements<A.Run>().Select(r => new PowerPointTextRun(r, _slidePart)).ToList();

        /// <summary>
        /// Adds a run to the paragraph.
        /// </summary>
        public PowerPointTextRun AddRun(string text, Action<PowerPointTextRun>? configure = null) {
            A.Run run = InsertRun(text);
            var wrapper = new PowerPointTextRun(run, _slidePart);
            configure?.Invoke(wrapper);
            return wrapper;
        }

        /// <summary>
        /// Gets or sets paragraph alignment.
        /// </summary>
        public A.TextAlignmentTypeValues? Alignment {
            get => Paragraph.ParagraphProperties?.Alignment?.Value;
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.Alignment = value;
            }
        }

        /// <summary>
        /// Sets paragraph alignment and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetAlignment(A.TextAlignmentTypeValues alignment) {
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
        /// Gets or sets paragraph indentation in points.
        /// </summary>
        public double? IndentPoints {
            get => FromTextCoordinate(Paragraph.ParagraphProperties?.Indent?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.Indent = ToTextCoordinate(value);
            }
        }

        /// <summary>
        /// Gets or sets paragraph left margin in points.
        /// </summary>
        public double? LeftMarginPoints {
            get => FromTextCoordinate(Paragraph.ParagraphProperties?.LeftMargin?.Value);
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.LeftMargin = ToTextCoordinate(value);
            }
        }

        /// <summary>
        /// Sets indentation in points and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetIndentPoints(double points) {
            IndentPoints = points;
            return this;
        }

        /// <summary>
        /// Sets indentation in centimeters and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetIndentCm(double centimeters) {
            return SetIndentPoints(PointsFromCentimeters(centimeters));
        }

        /// <summary>
        /// Sets indentation in inches and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetIndentInches(double inches) {
            return SetIndentPoints(PointsFromInches(inches));
        }

        /// <summary>
        /// Sets left margin in points and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetLeftMarginPoints(double points) {
            LeftMarginPoints = points;
            return this;
        }

        /// <summary>
        /// Sets left margin in centimeters and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetLeftMarginCm(double centimeters) {
            return SetLeftMarginPoints(PointsFromCentimeters(centimeters));
        }

        /// <summary>
        /// Sets left margin in inches and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetLeftMarginInches(double inches) {
            return SetLeftMarginPoints(PointsFromInches(inches));
        }

        /// <summary>
        /// Sets a hanging indent in points and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetHangingPoints(double points) {
            LeftMarginPoints = points;
            IndentPoints = -points;
            return this;
        }

        /// <summary>
        /// Sets a hanging indent in centimeters and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetHangingCm(double centimeters) {
            return SetHangingPoints(PointsFromCentimeters(centimeters));
        }

        /// <summary>
        /// Sets a hanging indent in inches and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetHangingInches(double inches) {
            return SetHangingPoints(PointsFromInches(inches));
        }

        /// <summary>
        /// Gets or sets line spacing in points.
        /// </summary>
        public double? LineSpacingPoints {
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
        /// Sets line spacing in points and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetLineSpacingPoints(double points) {
            LineSpacingPoints = points;
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
        /// Gets or sets space before the paragraph in points.
        /// </summary>
        public double? SpaceBeforePoints {
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
        /// Sets space before the paragraph in points and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetSpaceBeforePoints(double points) {
            SpaceBeforePoints = points;
            return this;
        }

        /// <summary>
        /// Gets or sets space after the paragraph in points.
        /// </summary>
        public double? SpaceAfterPoints {
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
        /// Sets space after the paragraph in points and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetSpaceAfterPoints(double points) {
            SpaceAfterPoints = points;
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
        /// Sets the current run font size in points and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetFontSize(int size) {
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
                    props.Append(new A.BulletFont { Typeface = value });
                }
            }
        }

        /// <summary>
        /// Gets or sets the bullet size in points.
        /// </summary>
        public int? BulletSizePoints {
            get {
                int? size = Paragraph.ParagraphProperties?.GetFirstChild<A.BulletSizePoints>()?.Val?.Value;
                return size != null ? size / 100 : null;
            }
            set {
                A.ParagraphProperties props = EnsureParagraphProperties();
                props.RemoveAllChildren<A.BulletSizePoints>();
                props.RemoveAllChildren<A.BulletSizePercentage>();
                if (value != null) {
                    props.Append(new A.BulletSizePoints { Val = value.Value * 100 });
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
                    props.Append(new A.BulletSizePercentage { Val = value.Value * 1000 });
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
        /// Sets the bullet size in points and returns the paragraph for chaining.
        /// </summary>
        public PowerPointParagraph SetBulletSizePoints(int sizePoints) {
            BulletSizePoints = sizePoints;
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
            props.Append(new A.CharacterBullet { Char = bulletChar.ToString() });
        }

        /// <summary>
        /// Applies an auto-numbered bullet to the paragraph.
        /// </summary>
        public void SetNumbered(A.TextAutoNumberSchemeValues style, int startAt = 1) {
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            props.Append(new A.AutoNumberedBullet { Type = style, StartAt = startAt });
        }

        /// <summary>
        /// Applies an auto-numbered bullet to the paragraph without resetting the start value.
        /// </summary>
        public void SetNumbered(A.TextAutoNumberSchemeValues style) {
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            props.Append(new A.AutoNumberedBullet { Type = style });
        }

        /// <summary>
        /// Applies a default auto-numbered bullet (Arabic period) to the paragraph.
        /// </summary>
        public void SetNumbered(int startAt = 1) {
            SetNumbered(A.TextAutoNumberSchemeValues.ArabicPeriod, startAt);
        }

        /// <summary>
        /// Clears any bullet/numbering from the paragraph.
        /// </summary>
        public void ClearBullet() {
            A.ParagraphProperties props = EnsureParagraphProperties();
            ClearBulletInternal(props);
            props.Append(new A.NoBullet());
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

        private static int? ToTextCoordinate(double? points) {
            return points != null ? (int)Math.Round(points.Value * 100) : null;
        }

        private static double? FromTextCoordinate(int? value) {
            return value != null ? value.Value / 100d : null;
        }

        private static int ToSpacingPoints(double points) {
            return (int)Math.Round(points * 100);
        }

        private static double? FromSpacingPoints(int? value) {
            return value != null ? value.Value / 100d : null;
        }

        private static int ToSpacingPercent(double multiplier) {
            return (int)Math.Round(multiplier * 100000d);
        }

        private static double? FromSpacingPercent(int? value) {
            return value != null ? value.Value / 100000d : null;
        }

        private static double PointsFromCentimeters(double centimeters) {
            long emus = PowerPointUnits.FromCentimeters(centimeters);
            return PowerPointUnits.ToPoints(emus);
        }

        private static double PointsFromInches(double inches) {
            long emus = PowerPointUnits.FromInches(inches);
            return PowerPointUnits.ToPoints(emus);
        }
    }
}
