using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.IWork;

namespace OfficeIMO.PowerPoint.IWork;

public static partial class PowerPointIWorkConverter {
    private static void AddRichTextBox(PowerPointSlide slide, IWorkTextBox source,
        double fallbackLeft, double fallbackTop, double fallbackWidth, double fallbackHeight,
        CancellationToken cancellationToken) {
        double left = source.Geometry?.LeftPoints ?? fallbackLeft;
        double top = source.Geometry?.TopPoints ?? fallbackTop;
        double width = source.Geometry?.WidthPoints ?? fallbackWidth;
        double height = source.Geometry?.HeightPoints ?? fallbackHeight;
        PowerPointTextBox textBox = slide.AddTextBoxPoints(string.Empty, left, top, width, height);
        textBox.Rotation = source.Geometry?.RotationDegrees;
        textBox.AltText = source.AccessibilityDescription;
        if (source.Hyperlink != null
            && Uri.TryCreate(source.Hyperlink, UriKind.RelativeOrAbsolute, out Uri? shapeLink)) {
            textBox.SetHyperlink(shapeLink);
        }
        textBox.Clear();
        bool first = true;
        var listState = new IWorkPowerPointListState();
        foreach (IWorkTextParagraph sourceParagraph in source.Content.Paragraphs) {
            cancellationToken.ThrowIfCancellationRequested();
            PowerPointParagraph paragraph;
            if (first) {
                paragraph = textBox.Paragraphs[0];
                paragraph.Text = string.Empty;
                first = false;
            } else {
                paragraph = textBox.AddParagraph();
            }
            ApplyParagraphStyle(paragraph, sourceParagraph,
                listState.StartsAtSourceLabel(sourceParagraph));
            WriteParagraphContent(paragraph, sourceParagraph, cancellationToken);
        }
    }

    private static void SetRichPresenterNotes(PowerPointNotes notes, IWorkTextContent source,
        CancellationToken cancellationToken) {
        IReadOnlyList<PowerPointParagraph> paragraphs = notes.SetParagraphs(
            source.Paragraphs.Select(_ => string.Empty));
        var listState = new IWorkPowerPointListState();
        for (int paragraphIndex = 0; paragraphIndex < source.Paragraphs.Count; paragraphIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            IWorkTextParagraph sourceParagraph = source.Paragraphs[paragraphIndex];
            PowerPointParagraph paragraph = paragraphs[paragraphIndex];
            ApplyParagraphStyle(paragraph, sourceParagraph,
                listState.StartsAtSourceLabel(sourceParagraph));
            WriteParagraphContent(paragraph, sourceParagraph, cancellationToken);
        }
        notes.Save();
    }

    private static void ApplyParagraphStyle(PowerPointParagraph paragraph,
        IWorkTextParagraph source, bool startsAtSourceLabel) {
        ApplyParagraphStyle(paragraph, source.Style, source.Text);
        if (source.ListLevel >= 0) {
            paragraph.Level = Math.Min(8, source.ListLevel);
            if (string.IsNullOrEmpty(source.ListLabel)) paragraph.SetBullet('\u2022');
            else if (TryParseNumbering(source.ListLabel!, out PowerPointNumberingScheme scheme,
                         out int start)) {
                if (startsAtSourceLabel) paragraph.SetNumbered(scheme, start);
                else paragraph.SetNumbered(scheme);
            } else if (source.ListLabel!.Length == 1) paragraph.SetBullet(source.ListLabel[0]);
        }
    }

    private static void ApplyParagraphStyle(PowerPointParagraph paragraph, IWorkParagraphStyle style, string text) {
        bool rightToLeft = OfficeTextElements.ResolveBaseDirection(text)
            == OfficeTextDirection.RightToLeft;
        paragraph.RightToLeft = rightToLeft;
        if (style.Alignment.HasValue) {
            paragraph.Alignment = style.Alignment.Value switch {
                IWorkTextAlignment.Natural => rightToLeft
                    ? PowerPointTextAlignment.Right
                    : PowerPointTextAlignment.Left,
                IWorkTextAlignment.Center => PowerPointTextAlignment.Center,
                IWorkTextAlignment.Right => PowerPointTextAlignment.Right,
                IWorkTextAlignment.Justified => PowerPointTextAlignment.Justified,
                _ => PowerPointTextAlignment.Left
            };
        }
        if (style.FirstLineIndentPoints.HasValue) paragraph.IndentPoints = style.FirstLineIndentPoints;
        if (style.LeftIndentPoints.HasValue) paragraph.LeftMarginPoints = style.LeftIndentPoints;
        if (style.SpaceBeforePoints.HasValue) paragraph.SpaceBeforePoints = style.SpaceBeforePoints;
        if (style.SpaceAfterPoints.HasValue) paragraph.SpaceAfterPoints = style.SpaceAfterPoints;
    }

    private sealed class IWorkPowerPointListState {
        private readonly HashSet<int> _observedLevels = new();
        private bool _inList;
        private ulong? _listIdentifier;

        internal bool StartsAtSourceLabel(IWorkTextParagraph paragraph) {
            if (paragraph.ListLevel < 0) {
                _inList = false;
                _listIdentifier = null;
                _observedLevels.Clear();
                return false;
            }
            if (!_inList || paragraph.ListIdentifier != _listIdentifier) {
                _inList = true;
                _listIdentifier = paragraph.ListIdentifier;
                _observedLevels.Clear();
            }
            return _observedLevels.Add(paragraph.ListLevel);
        }
    }

    private static void WriteParagraphContent(PowerPointParagraph paragraph,
        IWorkTextParagraph source, CancellationToken cancellationToken, IWorkTextStyle? defaultStyle = null, bool header = false) {
        paragraph.Text = string.Empty;
        foreach (PowerPointTextRun run in paragraph.Runs) {
            if (header) run.Bold = true;
            if (defaultStyle != null) ApplyTextStyle(run, defaultStyle);
            ApplyTextStyle(run, source.Style.TextStyle);
        }
        bool canReuseInitialRun = true;
        foreach (IWorkTextRun sourceRun in source.Runs) {
            cancellationToken.ThrowIfCancellationRequested();
            string[] lines = sourceRun.Text.Split(new[] { '\n' });
            for (int lineIndex = 0; lineIndex < lines.Length; lineIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                if (lineIndex > 0) {
                    paragraph.AddLineBreak();
                    canReuseInitialRun = false;
                }
                if (lines[lineIndex].Length > 0) {
                    AppendStyledText(paragraph, lines[lineIndex], sourceRun.Style,
                        sourceRun.Hyperlink, ref canReuseInitialRun, defaultStyle, header);
                }
            }
        }
    }

    private static bool TryParseNumbering(string label,
        out PowerPointNumberingScheme scheme, out int start) {
        const int MaximumStart = 32_767;
        scheme = PowerPointNumberingScheme.ArabicPeriod;
        start = 1;
        string marker = label.Trim();
        bool parenthesized = marker.Length > 2
            && marker[0] == '(' && marker[marker.Length - 1] == ')';
        bool rightParenthesis = !parenthesized && marker.EndsWith(")", StringComparison.Ordinal);
        bool period = !parenthesized && marker.EndsWith(".", StringComparison.Ordinal);
        string token = parenthesized
            ? marker.Substring(1, marker.Length - 2)
            : rightParenthesis || period ? marker.Substring(0, marker.Length - 1) : marker;
        if (token.Length == 0) return false;

        if (token.All(character => character is >= '0' and <= '9')) {
            if (!int.TryParse(token, System.Globalization.NumberStyles.None,
                    System.Globalization.CultureInfo.InvariantCulture, out start)
                || start is < 1 or > MaximumStart) return false;
            scheme = parenthesized ? PowerPointNumberingScheme.ArabicParenBoth
                : rightParenthesis ? PowerPointNumberingScheme.ArabicParenR
                : period ? PowerPointNumberingScheme.ArabicPeriod
                : PowerPointNumberingScheme.ArabicPlain;
            return true;
        }

        bool roman = token.All(character => "ivxlcdmIVXLCDM".IndexOf(character) >= 0)
            && (token.Length > 1 || "ivxIVX".IndexOf(token[0]) >= 0);
        if (roman) {
            if (!TryParseRoman(token, out start) || start > MaximumStart) return false;
            bool upper = token.All(character => character is >= 'A' and <= 'Z');
            if (!parenthesized && !rightParenthesis && !period) return false;
            scheme = upper
                ? parenthesized ? PowerPointNumberingScheme.RomanUpperCharacterParenBoth
                    : rightParenthesis ? PowerPointNumberingScheme.RomanUpperCharacterParenR
                    : PowerPointNumberingScheme.RomanUpperCharacterPeriod
                : parenthesized ? PowerPointNumberingScheme.RomanLowerCharacterParenBoth
                    : rightParenthesis ? PowerPointNumberingScheme.RomanLowerCharacterParenR
                    : PowerPointNumberingScheme.RomanLowerCharacterPeriod;
            return true;
        }

        bool uppercase = token.All(character => character is >= 'A' and <= 'Z');
        bool lowercase = token.All(character => character is >= 'a' and <= 'z');
        if (!uppercase && !lowercase
            || !TryParseAlphabetic(token, out start) || start > MaximumStart
            || !parenthesized && !rightParenthesis && !period) return false;
        scheme = uppercase
            ? parenthesized ? PowerPointNumberingScheme.AlphaUpperCharacterParenBoth
                : rightParenthesis ? PowerPointNumberingScheme.AlphaUpperCharacterParenR
                : PowerPointNumberingScheme.AlphaUpperCharacterPeriod
            : parenthesized ? PowerPointNumberingScheme.AlphaLowerCharacterParenBoth
                : rightParenthesis ? PowerPointNumberingScheme.AlphaLowerCharacterParenR
                : PowerPointNumberingScheme.AlphaLowerCharacterPeriod;
        return true;
    }

    private static bool TryParseAlphabetic(string token, out int value) {
        value = 0;
        foreach (char character in token) {
            int digit = char.ToUpperInvariant(character) - 'A' + 1;
            if (digit < 1 || digit > 26 || value > (int.MaxValue - digit) / 26) return false;
            value = value * 26 + digit;
        }
        return value > 0;
    }

    private static bool TryParseRoman(string token, out int value) {
        value = 0;
        int previous = 0;
        for (int index = token.Length - 1; index >= 0; index--) {
            int current = char.ToUpperInvariant(token[index]) switch {
                'I' => 1, 'V' => 5, 'X' => 10, 'L' => 50,
                'C' => 100, 'D' => 500, 'M' => 1000, _ => 0
            };
            if (current == 0) return false;
            int delta = current < previous ? -current : current;
            if (delta > 0 && value > int.MaxValue - delta
                || delta < 0 && value < int.MinValue - delta) return false;
            value += delta;
            if (current > previous) previous = current;
        }
        if (value <= 0) return false;
        bool upper = token.All(character => character is >= 'A' and <= 'Z');
        string canonical = FormatRoman(value);
        return string.Equals(token, upper ? canonical : canonical.ToLowerInvariant(),
            StringComparison.Ordinal);
    }

    private static string FormatRoman(int value) {
        var builder = new System.Text.StringBuilder();
        foreach ((int Number, string Token) part in new[] {
                     (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                     (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                     (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
                 }) {
            while (value >= part.Number) {
                builder.Append(part.Token);
                value -= part.Number;
            }
        }
        return builder.ToString();
    }

    private static void AppendStyledText(PowerPointParagraph paragraph, string text,
        IWorkTextStyle style, string? hyperlink, ref bool canReuseInitialRun, IWorkTextStyle? defaultStyle = null, bool header = false) {
        PowerPointTextRun run;
        if (canReuseInitialRun) {
            run = paragraph.Runs[0];
            run.Text = text;
            canReuseInitialRun = false;
        } else {
            run = paragraph.AddRun(text);
        }
        if (header) run.Bold = true;
        if (defaultStyle != null) ApplyTextStyle(run, defaultStyle);
        ApplyTextStyle(run, style);
        if (hyperlink != null
            && Uri.TryCreate(hyperlink, UriKind.RelativeOrAbsolute, out Uri? runLink)) {
            run.Hyperlink = runLink;
        }
    }

    private static void ApplyTextStyle(PowerPointTextRun run, IWorkTextStyle style) {
        if (style.Bold.HasValue) run.Bold = style.Bold.Value;
        if (style.Italic.HasValue) run.Italic = style.Italic.Value;
        if (style.Underline.HasValue) run.Underline = style.Underline.Value;
        if (style.Strikethrough.HasValue) run.Strikethrough = style.Strikethrough.Value;
        if (style.FontSizePoints.HasValue) run.FontSizePoints = style.FontSizePoints.Value;
        if (!string.IsNullOrWhiteSpace(style.FontName)) run.FontName = style.FontName;
        if (style.Color != null) run.Color = style.Color.RgbHex;
        if (style.BackgroundColor != null) run.HighlightColor = style.BackgroundColor.RgbHex;
    }
}
