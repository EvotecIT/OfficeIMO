using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Represents a textbox shape.
    /// </summary>
    public class PowerPointTextBox : PowerPointShape {
        private readonly SlidePart? _slidePart;

        internal PowerPointTextBox(Shape shape, SlidePart? slidePart = null) : base(shape) {
            _slidePart = slidePart;
        }

        private Shape Shape => (Shape)Element;

        private IEnumerable<A.Run> Runs => EnsureTextBody().Elements<A.Paragraph>().SelectMany(p => p.Elements<A.Run>());

        /// <summary>
        ///     Paragraphs contained in the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> Paragraphs =>
            EnsureTextBody().Elements<A.Paragraph>().Select(p => new PowerPointParagraph(p, _slidePart)).ToList();

        /// <summary>
        ///     Adds a paragraph to the textbox.
        /// </summary>
        public PowerPointParagraph AddParagraph(string text = "", Action<PowerPointParagraph>? configure = null,
            Action<PowerPointTextRun>? run = null) {
            TextBody textBody = EnsureTextBody();

            A.Paragraph paragraph = new();
            A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
            if (templateParagraph?.ParagraphProperties != null) {
                paragraph.ParagraphProperties = (A.ParagraphProperties)templateParagraph.ParagraphProperties.CloneNode(true);
            }

            A.Run newRun = new(new A.Text(text ?? string.Empty));
            A.RunProperties? templateRunProps = templateParagraph?
                .Elements<A.Run>()
                .Select(r => r.RunProperties)
                .FirstOrDefault(rp => rp != null);
            if (templateRunProps != null) {
                newRun.RunProperties = (A.RunProperties)templateRunProps.CloneNode(true);
            }
            paragraph.Append(newRun);

            A.EndParagraphRunProperties? templateEnd = templateParagraph?.GetFirstChild<A.EndParagraphRunProperties>();
            if (templateEnd != null) {
                paragraph.Append((A.EndParagraphRunProperties)templateEnd.CloneNode(true));
            }

            textBody.Append(paragraph);

            var wrapper = new PowerPointParagraph(paragraph, _slidePart);
            configure?.Invoke(wrapper);
            if (run != null) {
                var runWrapper = wrapper.Runs.FirstOrDefault() ?? wrapper.AddRun(text ?? string.Empty);
                run.Invoke(runWrapper);
            }

            return wrapper;
        }

        /// <summary>
        ///     Adds multiple paragraphs to the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddParagraphs(IEnumerable<string> paragraphs,
            Action<PowerPointParagraph>? configure = null) {
            if (paragraphs == null) {
                throw new ArgumentNullException(nameof(paragraphs));
            }

            var results = new List<PowerPointParagraph>();
            foreach (string paragraphText in paragraphs) {
                PowerPointParagraph paragraph = AddParagraph(paragraphText ?? string.Empty);
                configure?.Invoke(paragraph);
                results.Add(paragraph);
            }

            return results;
        }

        /// <summary>
        ///     Replaces all paragraphs with the provided content.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetParagraphs(IEnumerable<string> paragraphs,
            Action<PowerPointParagraph>? configure = null) {
            if (paragraphs == null) {
                throw new ArgumentNullException(nameof(paragraphs));
            }

            Clear();
            return AddParagraphs(paragraphs, configure);
        }

        /// <summary>
        ///     Adds a bulleted list to the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddBullets(IEnumerable<string> bullets, int level = 0,
            char bulletChar = '\u2022', Action<PowerPointParagraph>? configure = null) {
            if (bullets == null) {
                throw new ArgumentNullException(nameof(bullets));
            }

            var results = new List<PowerPointParagraph>();
            foreach (string bullet in bullets) {
                PowerPointParagraph paragraph = AddParagraph(bullet ?? string.Empty);
                paragraph.SetBullet(bulletChar);
                if (level > 0) {
                    paragraph.Level = level;
                }
                ApplyDefaultListSpacing(paragraph);
                configure?.Invoke(paragraph);
                results.Add(paragraph);
            }

            return results;
        }

        /// <summary>
        ///     Replaces all paragraphs with a bulleted list.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetBullets(IEnumerable<string> bullets, int level = 0,
            char bulletChar = '\u2022', Action<PowerPointParagraph>? configure = null) {
            if (bullets == null) {
                throw new ArgumentNullException(nameof(bullets));
            }

            Clear();
            return AddBullets(bullets, level, bulletChar, configure);
        }

        /// <summary>
        ///     Adds a numbered list to the textbox.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddNumberedList(IEnumerable<string> items,
            A.TextAutoNumberSchemeValues style, int startAt = 1,
            int level = 0, Action<PowerPointParagraph>? configure = null) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            var results = new List<PowerPointParagraph>();
            bool first = true;
            foreach (string item in items) {
                PowerPointParagraph paragraph = AddParagraph(item ?? string.Empty);
                if (first) {
                    paragraph.SetNumbered(style, startAt);
                    first = false;
                } else {
                    paragraph.SetNumbered(style);
                }
                if (level > 0) {
                    paragraph.Level = level;
                }
                ApplyDefaultListSpacing(paragraph);
                configure?.Invoke(paragraph);
                results.Add(paragraph);
            }

            return results;
        }

        /// <summary>
        ///     Adds a numbered list to the textbox using the default numbering style.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> AddNumberedList(IEnumerable<string> items,
            int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            return AddNumberedList(items, A.TextAutoNumberSchemeValues.ArabicPeriod, startAt, level, configure);
        }

        /// <summary>
        ///     Replaces all paragraphs with a numbered list.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetNumberedList(IEnumerable<string> items,
            A.TextAutoNumberSchemeValues style, int startAt = 1,
            int level = 0, Action<PowerPointParagraph>? configure = null) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            Clear();
            return AddNumberedList(items, style, startAt, level, configure);
        }

        /// <summary>
        ///     Replaces all paragraphs with a numbered list using the default numbering style.
        /// </summary>
        public IReadOnlyList<PowerPointParagraph> SetNumberedList(IEnumerable<string> items,
            int startAt = 1, int level = 0, Action<PowerPointParagraph>? configure = null) {
            return SetNumberedList(items, A.TextAutoNumberSchemeValues.ArabicPeriod, startAt, level, configure);
        }

        /// <summary>
        ///     Applies a text style to the textbox.
        /// </summary>
        public PowerPointTextBox ApplyTextStyle(PowerPointTextStyle style, bool applyToRuns = false) {
            var paragraphs = GetParagraphsForStyling();
            foreach (PowerPointParagraph paragraph in paragraphs) {
                if (applyToRuns) {
                    IReadOnlyList<PowerPointTextRun> runs = paragraph.Runs;
                    if (runs.Count == 0) {
                        runs = new List<PowerPointTextRun> { paragraph.AddRun(string.Empty) };
                    }
                    foreach (PowerPointTextRun run in runs) {
                        style.Apply(run);
                    }
                } else {
                    style.Apply(paragraph);
                }
            }

            return this;
        }

        /// <summary>
        ///     Applies a paragraph style to the textbox.
        /// </summary>
        public PowerPointTextBox ApplyParagraphStyle(PowerPointParagraphStyle style) {
            var paragraphs = GetParagraphsForStyling();
            foreach (PowerPointParagraph paragraph in paragraphs) {
                style.Apply(paragraph);
            }

            return this;
        }

        /// <summary>
        ///     Applies automatic spacing defaults to the textbox paragraphs.
        /// </summary>
        public PowerPointTextBox ApplyAutoSpacing(double lineSpacingMultiplier = 1.2,
            double? spaceBeforePoints = null, double? spaceAfterPoints = null) {
            PowerPointParagraphStyle style = new(lineSpacingMultiplier: lineSpacingMultiplier,
                spaceBeforePoints: spaceBeforePoints,
                spaceAfterPoints: spaceAfterPoints);
            return ApplyParagraphStyle(style);
        }

        /// <summary>
        ///     Replaces text within the textbox while preserving run formatting.
        /// </summary>
        public int ReplaceText(string oldValue, string newValue) {
            int count = 0;
            foreach (PowerPointParagraph paragraph in Paragraphs) {
                count += paragraph.ReplaceText(oldValue, newValue);
            }
            return count;
        }

        /// <summary>
        ///     Removes all paragraphs from the textbox.
        /// </summary>
        public void Clear() {
            EnsureTextBody().RemoveAllChildren<A.Paragraph>();
        }

        /// <summary>
        ///     Text contained in the textbox.
        /// </summary>
        public string Text {
            get {
                var paragraphs = EnsureTextBody().Elements<A.Paragraph>()
                    .Select(p => p.InnerText ?? string.Empty)
                    .ToList();
                return paragraphs.Count == 0 ? string.Empty : string.Join(Environment.NewLine, paragraphs);
            }
            set {
                string textValue = value ?? string.Empty;
                TextBody textBody = EnsureTextBody();

                A.Paragraph? templateParagraph = textBody.Elements<A.Paragraph>().FirstOrDefault();
                A.ParagraphProperties? templateParagraphProperties = templateParagraph?.GetFirstChild<A.ParagraphProperties>();
                A.EndParagraphRunProperties? templateEndParagraphRunProperties = templateParagraph?.GetFirstChild<A.EndParagraphRunProperties>();
                A.RunProperties? templateRunProperties = templateParagraph?
                    .Elements<A.Run>()
                    .Select(r => r.RunProperties)
                    .FirstOrDefault(rp => rp != null);

                textBody.RemoveAllChildren<A.Paragraph>();

                string[] lines = textValue.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                foreach (string line in lines) {
                    A.Paragraph paragraph = new();
                    if (templateParagraphProperties != null) {
                        paragraph.Append((A.ParagraphProperties)templateParagraphProperties.CloneNode(true));
                    }

                    A.Run run = new();
                    if (templateRunProperties != null) {
                        run.RunProperties = (A.RunProperties)templateRunProperties.CloneNode(true);
                    }

                    run.Append(new A.Text(line));
                    paragraph.Append(run);

                    if (templateEndParagraphRunProperties != null) {
                        paragraph.Append((A.EndParagraphRunProperties)templateEndParagraphRunProperties.CloneNode(true));
                    }

                    textBody.Append(paragraph);
                }
            }
        }

        /// <summary>
        ///     Returns true if the textbox is tied to a slide placeholder.
        /// </summary>
        public bool IsPlaceholder {
            get {
                PlaceholderShape? placeholder = GetPlaceholderShape();
                return placeholder?.Type != null || placeholder?.Index != null;
            }
        }

        /// <summary>
        ///     Gets or sets the placeholder type for this textbox.
        /// </summary>
        public PlaceholderValues? PlaceholderType {
            get => GetPlaceholderShape()?.Type?.Value;
            set {
                if (value == null) {
                    PlaceholderShape? placeholder = GetPlaceholderShape();
                    if (placeholder != null) {
                        placeholder.Type = null;
                        if (placeholder.Index == null && placeholder.HasChildren == false) {
                            placeholder.Remove();
                        }
                    }
                    return;
                }

                PlaceholderShape shape = EnsurePlaceholderShape();
                shape.Type = value.Value;
            }
        }

        /// <summary>
        ///     Gets or sets the placeholder index for this textbox.
        /// </summary>
        public uint? PlaceholderIndex {
            get => GetPlaceholderShape()?.Index?.Value;
            set {
                if (value == null) {
                    PlaceholderShape? placeholder = GetPlaceholderShape();
                    if (placeholder != null) {
                        placeholder.Index = null;
                        if (placeholder.Type == null && placeholder.HasChildren == false) {
                            placeholder.Remove();
                        }
                    }
                    return;
                }

                PlaceholderShape shape = EnsurePlaceholderShape();
                shape.Index = value.Value;
            }
        }

        /// <summary>
        ///     Gets or sets the left text margin in points.
        /// </summary>
        public double? TextMarginLeftPoints {
            get => FromEmusInt(GetBodyProperties()?.LeftInset?.Value, PowerPointUnits.EmusPerPoint);
            set => EnsureBodyProperties().LeftInset = ToEmusInt(value, PowerPointUnits.EmusPerPoint);
        }

        /// <summary>
        ///     Gets or sets the right text margin in points.
        /// </summary>
        public double? TextMarginRightPoints {
            get => FromEmusInt(GetBodyProperties()?.RightInset?.Value, PowerPointUnits.EmusPerPoint);
            set => EnsureBodyProperties().RightInset = ToEmusInt(value, PowerPointUnits.EmusPerPoint);
        }

        /// <summary>
        ///     Gets or sets the top text margin in points.
        /// </summary>
        public double? TextMarginTopPoints {
            get => FromEmusInt(GetBodyProperties()?.TopInset?.Value, PowerPointUnits.EmusPerPoint);
            set => EnsureBodyProperties().TopInset = ToEmusInt(value, PowerPointUnits.EmusPerPoint);
        }

        /// <summary>
        ///     Gets or sets the bottom text margin in points.
        /// </summary>
        public double? TextMarginBottomPoints {
            get => FromEmusInt(GetBodyProperties()?.BottomInset?.Value, PowerPointUnits.EmusPerPoint);
            set => EnsureBodyProperties().BottomInset = ToEmusInt(value, PowerPointUnits.EmusPerPoint);
        }

        /// <summary>
        ///     Gets or sets the left text margin in centimeters.
        /// </summary>
        public double? TextMarginLeftCm {
            get => FromEmusInt(GetBodyProperties()?.LeftInset?.Value, PowerPointUnits.EmusPerCentimeter);
            set => EnsureBodyProperties().LeftInset = ToEmusInt(value, PowerPointUnits.EmusPerCentimeter);
        }

        /// <summary>
        ///     Gets or sets the right text margin in centimeters.
        /// </summary>
        public double? TextMarginRightCm {
            get => FromEmusInt(GetBodyProperties()?.RightInset?.Value, PowerPointUnits.EmusPerCentimeter);
            set => EnsureBodyProperties().RightInset = ToEmusInt(value, PowerPointUnits.EmusPerCentimeter);
        }

        /// <summary>
        ///     Gets or sets the top text margin in centimeters.
        /// </summary>
        public double? TextMarginTopCm {
            get => FromEmusInt(GetBodyProperties()?.TopInset?.Value, PowerPointUnits.EmusPerCentimeter);
            set => EnsureBodyProperties().TopInset = ToEmusInt(value, PowerPointUnits.EmusPerCentimeter);
        }

        /// <summary>
        ///     Gets or sets the bottom text margin in centimeters.
        /// </summary>
        public double? TextMarginBottomCm {
            get => FromEmusInt(GetBodyProperties()?.BottomInset?.Value, PowerPointUnits.EmusPerCentimeter);
            set => EnsureBodyProperties().BottomInset = ToEmusInt(value, PowerPointUnits.EmusPerCentimeter);
        }

        /// <summary>
        ///     Gets or sets the left text margin in inches.
        /// </summary>
        public double? TextMarginLeftInches {
            get => FromEmusInt(GetBodyProperties()?.LeftInset?.Value, PowerPointUnits.EmusPerInch);
            set => EnsureBodyProperties().LeftInset = ToEmusInt(value, PowerPointUnits.EmusPerInch);
        }

        /// <summary>
        ///     Gets or sets the right text margin in inches.
        /// </summary>
        public double? TextMarginRightInches {
            get => FromEmusInt(GetBodyProperties()?.RightInset?.Value, PowerPointUnits.EmusPerInch);
            set => EnsureBodyProperties().RightInset = ToEmusInt(value, PowerPointUnits.EmusPerInch);
        }

        /// <summary>
        ///     Gets or sets the top text margin in inches.
        /// </summary>
        public double? TextMarginTopInches {
            get => FromEmusInt(GetBodyProperties()?.TopInset?.Value, PowerPointUnits.EmusPerInch);
            set => EnsureBodyProperties().TopInset = ToEmusInt(value, PowerPointUnits.EmusPerInch);
        }

        /// <summary>
        ///     Gets or sets the bottom text margin in inches.
        /// </summary>
        public double? TextMarginBottomInches {
            get => FromEmusInt(GetBodyProperties()?.BottomInset?.Value, PowerPointUnits.EmusPerInch);
            set => EnsureBodyProperties().BottomInset = ToEmusInt(value, PowerPointUnits.EmusPerInch);
        }

        /// <summary>
        ///     Sets all text margins in points.
        /// </summary>
        public PowerPointTextBox SetTextMarginsPoints(double left, double top, double right, double bottom) {
            TextMarginLeftPoints = left;
            TextMarginTopPoints = top;
            TextMarginRightPoints = right;
            TextMarginBottomPoints = bottom;
            return this;
        }

        /// <summary>
        ///     Sets all text margins in centimeters.
        /// </summary>
        public PowerPointTextBox SetTextMarginsCm(double leftCm, double topCm, double rightCm, double bottomCm) {
            TextMarginLeftCm = leftCm;
            TextMarginTopCm = topCm;
            TextMarginRightCm = rightCm;
            TextMarginBottomCm = bottomCm;
            return this;
        }

        /// <summary>
        ///     Sets all text margins in inches.
        /// </summary>
        public PowerPointTextBox SetTextMarginsInches(double leftInches, double topInches, double rightInches, double bottomInches) {
            TextMarginLeftInches = leftInches;
            TextMarginTopInches = topInches;
            TextMarginRightInches = rightInches;
            TextMarginBottomInches = bottomInches;
            return this;
        }

        /// <summary>
        ///     Gets or sets the vertical anchoring of text inside the textbox.
        /// </summary>
        public A.TextAnchoringTypeValues? TextVerticalAlignment {
            get => GetBodyProperties()?.Anchor?.Value;
            set => EnsureBodyProperties().Anchor = value;
        }

        /// <summary>
        ///     Gets or sets the text direction (horizontal/vertical).
        /// </summary>
        public A.TextVerticalValues? TextDirection {
            get => GetBodyProperties()?.Vertical?.Value;
            set => EnsureBodyProperties().Vertical = value;
        }

        /// <summary>
        ///     Gets or sets the text auto-fit behavior.
        /// </summary>
        public PowerPointTextAutoFit? TextAutoFit {
            get {
                A.BodyProperties? body = GetBodyProperties();
                if (body == null) {
                    return null;
                }
                if (body.GetFirstChild<A.NoAutoFit>() != null) {
                    return PowerPointTextAutoFit.None;
                }
                if (body.GetFirstChild<A.NormalAutoFit>() != null) {
                    return PowerPointTextAutoFit.Normal;
                }
                if (body.GetFirstChild<A.ShapeAutoFit>() != null) {
                    return PowerPointTextAutoFit.Shape;
                }
                return null;
            }
            set {
                A.BodyProperties body = EnsureBodyProperties();
                body.RemoveAllChildren<A.NoAutoFit>();
                body.RemoveAllChildren<A.NormalAutoFit>();
                body.RemoveAllChildren<A.ShapeAutoFit>();

                if (value == null) {
                    return;
                }

                switch (value.Value) {
                    case PowerPointTextAutoFit.None:
                        body.Append(new A.NoAutoFit());
                        break;
                    case PowerPointTextAutoFit.Normal:
                        body.Append(new A.NormalAutoFit());
                        break;
                    case PowerPointTextAutoFit.Shape:
                        body.Append(new A.ShapeAutoFit());
                        break;
                }
            }
        }

        /// <summary>
        ///     Gets or sets detailed auto-fit options (only applies to Normal auto-fit).
        /// </summary>
        public PowerPointTextAutoFitOptions? TextAutoFitOptions {
            get {
                A.BodyProperties? body = GetBodyProperties();
                A.NormalAutoFit? normal = body?.GetFirstChild<A.NormalAutoFit>();
                if (normal == null) {
                    return null;
                }
                return PowerPointTextAutoFitOptions.FromOpenXmlValues(
                    normal.FontScale?.Value, normal.LineSpaceReduction?.Value);
            }
            set {
                if (value == null) {
                    A.BodyProperties? body = GetBodyProperties();
                    body?.RemoveAllChildren<A.NormalAutoFit>();
                    return;
                }

                A.BodyProperties bodyProperties = EnsureBodyProperties();
                bodyProperties.RemoveAllChildren<A.NoAutoFit>();
                bodyProperties.RemoveAllChildren<A.ShapeAutoFit>();
                A.NormalAutoFit normal = bodyProperties.GetFirstChild<A.NormalAutoFit>()
                    ?? bodyProperties.AppendChild(new A.NormalAutoFit());
                ApplyNormalAutoFitOptions(normal, value.Value);
            }
        }

        /// <summary>
        ///     Sets the auto-fit mode and optional Normal auto-fit options.
        /// </summary>
        public PowerPointTextBox SetTextAutoFit(PowerPointTextAutoFit fit, PowerPointTextAutoFitOptions? options = null) {
            TextAutoFit = fit;
            if (fit == PowerPointTextAutoFit.Normal && options != null) {
                A.BodyProperties body = EnsureBodyProperties();
                A.NormalAutoFit normal = body.GetFirstChild<A.NormalAutoFit>() ?? body.AppendChild(new A.NormalAutoFit());
                ApplyNormalAutoFitOptions(normal, options.Value);
            }
            return this;
        }

        /// <summary>
        ///     Gets or sets a value indicating whether the text is bold.       
        /// </summary>
        public bool Bold {
            get {
                A.Run? run = Runs.FirstOrDefault();
                return run?.RunProperties?.Bold?.Value == true;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    runProps.Bold = value ? true : null;
                }
            }
        }

        /// <summary>
        ///     Gets or sets a value indicating whether the text is italic.
        /// </summary>
        public bool Italic {
            get {
                A.Run? run = Runs.FirstOrDefault();
                return run?.RunProperties?.Italic?.Value == true;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    runProps.Italic = value ? true : null;
                }
            }
        }

        /// <summary>
        ///     Gets or sets the font size in points.
        /// </summary>
        public int? FontSize {
            get {
                A.Run? run = Runs.FirstOrDefault();
                int? size = run?.RunProperties?.FontSize?.Value;
                return size != null ? size / 100 : null;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    runProps.FontSize = value != null ? value * 100 : null;
                }
            }
        }

        /// <summary>
        ///     Gets or sets the font name.
        /// </summary>
        public string? FontName {
            get {
                A.Run? run = Runs.FirstOrDefault();
                return run?.RunProperties?.GetFirstChild<A.LatinFont>()?.Typeface;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    runProps.RemoveAllChildren<A.LatinFont>();
                    if (value != null) {
                        runProps.Append(new A.LatinFont { Typeface = value });
                    }
                }
            }
        }

        /// <summary>
        ///     Gets or sets the text color in hexadecimal format (e.g. "FF0000").
        /// </summary>
        public string? Color {
            get {
                A.Run? run = Runs.FirstOrDefault();
                return run?.RunProperties?.GetFirstChild<A.SolidFill>()?.RgbColorModelHex?.Val;
            }
            set {
                foreach (A.Run run in Runs) {
                    A.RunProperties runProps = run.RunProperties ??= new A.RunProperties();
                    // preserve existing fonts so we can reapply them after the fill (schema: fill must precede fonts)
                    var latin = runProps.GetFirstChild<A.LatinFont>();
                    var ea = runProps.GetFirstChild<A.EastAsianFont>();
                    var cs = runProps.GetFirstChild<A.ComplexScriptFont>();

                    runProps.RemoveAllChildren<A.SolidFill>();
                    runProps.RemoveAllChildren<A.LatinFont>();
                    runProps.RemoveAllChildren<A.EastAsianFont>();
                    runProps.RemoveAllChildren<A.ComplexScriptFont>();

                    if (value != null) {
                        runProps.Append(new A.SolidFill(new A.RgbColorModelHex { Val = value }));
                    }

                    if (latin != null) runProps.Append((A.LatinFont)latin.CloneNode(true));
                    if (ea != null) runProps.Append((A.EastAsianFont)ea.CloneNode(true));
                    if (cs != null) runProps.Append((A.ComplexScriptFont)cs.CloneNode(true));
                }
            }
        }

        private static void ApplyNormalAutoFitOptions(A.NormalAutoFit normal, PowerPointTextAutoFitOptions options) {
            normal.FontScale = options.FontScaleValue;
            normal.LineSpaceReduction = options.LineSpaceReductionValue;
        }

        /// <summary>
        ///     Adds a new bulleted paragraph to the textbox.
        /// </summary>
        public PowerPointParagraph AddBullet(string text) {
            PowerPointParagraph paragraph = AddParagraph(text);
            paragraph.SetBullet();
            ApplyDefaultListSpacing(paragraph);
            return paragraph;
        }

        /// <summary>
        ///     Adds a numbered item to the textbox.
        /// </summary>
        public PowerPointParagraph AddNumberedItem(string text, A.TextAutoNumberSchemeValues style, int startAt = 1) {
            PowerPointParagraph paragraph = AddParagraph(text);
            paragraph.SetNumbered(style, startAt);
            ApplyDefaultListSpacing(paragraph);
            return paragraph;
        }

        /// <summary>
        ///     Adds a numbered item to the textbox using the default numbering style.
        /// </summary>
        public PowerPointParagraph AddNumberedItem(string text, int startAt = 1) {
            PowerPointParagraph paragraph = AddParagraph(text);
            paragraph.SetNumbered(startAt);
            ApplyDefaultListSpacing(paragraph);
            return paragraph;
        }

        private TextBody EnsureTextBody() {
            TextBody? existingTextBody = Shape.TextBody;
            if (existingTextBody == null) {
                existingTextBody = new TextBody(new A.BodyProperties(), new A.ListStyle());
                Shape.TextBody = existingTextBody;
            }
            return existingTextBody;
        }

        private List<PowerPointParagraph> GetParagraphsForStyling() {
            var paragraphs = Paragraphs.ToList();
            if (paragraphs.Count == 0) {
                paragraphs.Add(AddParagraph());
            }
            return paragraphs;
        }

        private static readonly PowerPointParagraphStyle DefaultListParagraphStyle =
            new(lineSpacingMultiplier: 1.15, spaceAfterPoints: 2);

        private static void ApplyDefaultListSpacing(PowerPointParagraph paragraph) {
            DefaultListParagraphStyle.Apply(paragraph);
        }

        private A.BodyProperties? GetBodyProperties() {
            return Shape.TextBody?.GetFirstChild<A.BodyProperties>();
        }

        private A.BodyProperties EnsureBodyProperties() {
            TextBody textBody = EnsureTextBody();
            A.BodyProperties? body = textBody.GetFirstChild<A.BodyProperties>();
            if (body == null) {
                body = new A.BodyProperties();
                OpenXmlElement? first = textBody.FirstChild;
                if (first != null) {
                    textBody.InsertBefore(body, first);
                } else {
                    textBody.Append(body);
                }
            }
            return body;
        }

        private PlaceholderShape? GetPlaceholderShape() {
            return Shape.NonVisualShapeProperties?
                .ApplicationNonVisualDrawingProperties?
                .PlaceholderShape;
        }

        private PlaceholderShape EnsurePlaceholderShape() {
            NonVisualShapeProperties nonVisual = Shape.NonVisualShapeProperties ??= new NonVisualShapeProperties();
            ApplicationNonVisualDrawingProperties app = nonVisual.ApplicationNonVisualDrawingProperties ??
                                                       new ApplicationNonVisualDrawingProperties();
            nonVisual.ApplicationNonVisualDrawingProperties ??= app;
            return app.PlaceholderShape ??= new PlaceholderShape();
        }

        private static int? ToEmusInt(double? value, double emusPerUnit) {
            return value != null ? (int)Math.Round(value.Value * emusPerUnit) : null;
        }

        private static double? FromEmusInt(int? emus, double emusPerUnit) {
            return emus != null ? emus.Value / emusPerUnit : null;
        }
    }
}
