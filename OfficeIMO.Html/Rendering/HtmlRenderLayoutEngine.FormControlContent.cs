using System.Globalization;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private void AddFormControlContent(
        ICollection<HtmlRenderVisual> visuals,
        IElement element,
        HtmlRenderBoxStyle style,
        double boxX,
        double boxY,
        double boxWidth,
        double boxHeight,
        string source) {
        string tag = element.TagName.ToLowerInvariant();
        string type = NormalizeInputType(element);
        double contentX = boxX + style.BorderLeftWidth + style.PaddingLeft;
        double contentY = boxY + style.BorderTopWidth + style.PaddingTop;
        double contentWidth = Math.Max(0.01D, boxWidth - style.HorizontalInsets);
        double contentHeight = Math.Max(0.01D, boxHeight - style.VerticalInsets);

        if (tag == "input" && type == "checkbox") {
            if (HtmlFormControlSemantics.IsIndeterminate(element)) {
                OfficeShape mark = OfficeShape.Rectangle(contentWidth * 0.65D, Math.Max(2D, contentHeight * 0.18D));
                mark.FillColor = ControlAccentColor;
                mark.StrokeWidth = 0D;
                visuals.Add(new HtmlRenderShape(mark, contentX + contentWidth * 0.175D, contentY + contentHeight * 0.41D, visuals.Count, source: source + ":indeterminate"));
            } else if (HtmlFormControlSemantics.IsEffectivelyChecked(element)) AddCheckboxMark(visuals, contentX, contentY, contentWidth, contentHeight, source);
            return;
        }
        if (tag == "input" && type == "radio") {
            ReplaceControlBackgroundWithRadio(visuals, boxX, boxY, boxWidth, boxHeight, style, source);
            if (HtmlFormControlSemantics.IsEffectivelyChecked(element)) AddRadioMark(visuals, contentX, contentY, contentWidth, contentHeight, source);
            return;
        }
        if (tag == "input" && type == "range") {
            AddRangeContent(visuals, element, contentX, contentY, contentWidth, contentHeight, source);
            return;
        }
        if (tag == "input" && type == "color") {
            AddColorContent(visuals, element, contentX, contentY, contentWidth, contentHeight, source);
            return;
        }
        if (tag == "progress" || tag == "meter") {
            AddGaugeContent(visuals, element, tag, contentX, contentY, contentWidth, contentHeight, style, source);
            return;
        }

        if (tag == "textarea") {
            string text = NormalizeControlMultilineText(HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty);
            bool placeholder = text.Length == 0;
            if (placeholder) text = NormalizeControlMultilineText(element.GetAttribute("placeholder") ?? string.Empty);
            bool softWrap = !string.Equals(element.GetAttribute("wrap"), "off", StringComparison.OrdinalIgnoreCase);
            AddMultilineControlText(visuals, text, contentX, contentY, contentWidth, contentHeight, style, placeholder, source, softWrap);
            return;
        }

        if (tag == "select") {
            AddSelectContent(visuals, element, contentX, contentY, contentWidth, contentHeight, style, source);
            return;
        }

        string value;
        bool isPlaceholder = false;
        OfficeTextAlignment alignment = OfficeTextAlignment.Left;
        if (tag == "button" || tag == "input" && IsButtonInputType(type)) {
            value = ResolveButtonLabel(element, type);
            alignment = OfficeTextAlignment.Center;
        } else if (tag == "input" && type == "file") {
            value = "Choose file";
        } else {
            value = tag == "input"
                ? HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty
                : NormalizeControlText(element.GetAttribute("value"));
            if (type == "password" && value.Length > 0) value = new string('*', value.Length);
            if (value.Length == 0 && HtmlFormControlSemantics.IsPlaceholderApplicable(tag, type)) {
                value = NormalizeControlText(element.GetAttribute("placeholder"));
                isPlaceholder = value.Length > 0;
            }
        }

        AddSingleLineControlText(
            visuals,
            value,
            contentX,
            contentY,
            contentWidth,
            contentHeight,
            style,
            isPlaceholder,
            alignment,
            source);
    }

    private static void AddCheckboxMark(
        ICollection<HtmlRenderVisual> visuals,
        double x,
        double y,
        double width,
        double height,
        string source) {
        double left = x + width * 0.20D;
        double middleX = x + width * 0.43D;
        double middleY = y + height * 0.72D;
        OfficeShape first = OfficeShape.Line(left, y + height * 0.52D, middleX, middleY);
        first.StrokeColor = ControlAccentColor;
        first.StrokeWidth = Math.Max(1.5D, width * 0.13D);
        first.StrokeLineCap = OfficeStrokeLineCap.Round;
        visuals.Add(new HtmlRenderShape(first, Math.Min(left, middleX), Math.Min(y + height * 0.52D, middleY), visuals.Count, source: source + ":checked"));

        double right = x + width * 0.84D;
        OfficeShape second = OfficeShape.Line(middleX, middleY, right, y + height * 0.25D);
        second.StrokeColor = ControlAccentColor;
        second.StrokeWidth = first.StrokeWidth;
        second.StrokeLineCap = OfficeStrokeLineCap.Round;
        visuals.Add(new HtmlRenderShape(second, Math.Min(middleX, right), Math.Min(middleY, y + height * 0.25D), visuals.Count, source: source + ":checked"));
    }

    private static void ReplaceControlBackgroundWithRadio(
        ICollection<HtmlRenderVisual> visuals,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBoxStyle style,
        string source) {
        visuals.Clear();
        OfficeShape circle = OfficeShape.Ellipse(width, height);
        circle.FillColor = style.BackgroundColor;
        circle.StrokeColor = style.BorderColor;
        circle.StrokeWidth = Math.Max(1D, style.BorderWidth);
        visuals.Add(new HtmlRenderShape(circle, x, y, visuals.Count, source: source));
    }

    private static void AddRadioMark(
        ICollection<HtmlRenderVisual> visuals,
        double x,
        double y,
        double width,
        double height,
        string source) {
        double dotWidth = Math.Max(2D, width * 0.48D);
        double dotHeight = Math.Max(2D, height * 0.48D);
        OfficeShape dot = OfficeShape.Ellipse(dotWidth, dotHeight);
        dot.FillColor = ControlAccentColor;
        dot.StrokeWidth = 0D;
        visuals.Add(new HtmlRenderShape(
            dot,
            x + (width - dotWidth) / 2D,
            y + (height - dotHeight) / 2D,
            visuals.Count,
            source: source + ":checked"));
    }

    private static void AddRangeContent(
        ICollection<HtmlRenderVisual> visuals,
        IElement element,
        double x,
        double y,
        double width,
        double height,
        string source) {
        double fraction = HtmlFormControlSemantics.GetRangeFraction(element);
        double trackHeight = Math.Max(2D, Math.Min(4D, height * 0.25D));
        double trackY = y + (height - trackHeight) / 2D;
        OfficeShape track = OfficeShape.RoundedRectangle(width, trackHeight, trackHeight / 2D);
        track.FillColor = OfficeColor.FromRgb(196, 196, 196);
        track.StrokeWidth = 0D;
        visuals.Add(new HtmlRenderShape(track, x, trackY, visuals.Count, source: source + ":track"));

        double thumbSize = Math.Max(8D, Math.Min(height, 14D));
        OfficeShape thumb = OfficeShape.Ellipse(thumbSize, thumbSize);
        thumb.FillColor = ControlAccentColor;
        thumb.StrokeColor = OfficeColor.White;
        thumb.StrokeWidth = 1D;
        visuals.Add(new HtmlRenderShape(
            thumb,
            x + fraction * Math.Max(0D, width - thumbSize),
            y + (height - thumbSize) / 2D,
            visuals.Count,
            source: source + ":thumb"));
    }

    private static void AddColorContent(
        ICollection<HtmlRenderVisual> visuals,
        IElement element,
        double x,
        double y,
        double width,
        double height,
        string source) {
        string value = HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty;
        OfficeColor color = HtmlRenderCssValues.TryColor(value, out OfficeColor parsed)
            ? parsed
            : OfficeColor.Black;
        OfficeShape swatch = OfficeShape.Rectangle(width, height);
        swatch.FillColor = color;
        swatch.StrokeColor = OfficeColor.FromRgb(96, 96, 96);
        swatch.StrokeWidth = 1D;
        visuals.Add(new HtmlRenderShape(swatch, x, y, visuals.Count, source: source + ":swatch"));
    }

    private static void AddGaugeContent(
        ICollection<HtmlRenderVisual> visuals,
        IElement element,
        string tag,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBoxStyle style,
        string source) {
        double fraction = ResolveNumericFraction(element, 0D, tag == "progress" ? 1D : 1D, tag == "progress" ? 0D : 0D);
        OfficeShape track = OfficeShape.RoundedRectangle(width, height, Math.Min(3D, height / 2D));
        track.FillColor = OfficeColor.FromRgb(224, 224, 224);
        track.StrokeWidth = 0D;
        visuals.Add(new HtmlRenderShape(track, x, y, visuals.Count, source: source + ":track"));
        double fillWidth = Math.Max(0.01D, width * fraction);
        OfficeShape fill = OfficeShape.RoundedRectangle(fillWidth, height, Math.Min(3D, Math.Min(fillWidth, height) / 2D));
        fill.FillColor = tag == "meter" && fraction < 0.25D ? OfficeColor.FromRgb(206, 73, 52) : ControlAccentColor;
        fill.StrokeWidth = 0D;
        visuals.Add(new HtmlRenderShape(fill, x, y, visuals.Count, source: source + ":value"));

        string label = Math.Round(fraction * 100D, MidpointRounding.AwayFromZero).ToString(CultureInfo.InvariantCulture) + "%";
        AddSingleLineControlText(visuals, label, x, y, width, height, style, false, OfficeTextAlignment.Center, source + ":label");
    }

    private void AddSelectContent(
        ICollection<HtmlRenderVisual> visuals,
        IElement element,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBoxStyle style,
        string source) {
        bool listBox = element.HasAttribute("multiple")
            || HtmlFormControlSemantics.GetSelectDisplaySize(element) > 1;
        if (listBox) {
            string[] values = HtmlFormControlSemantics.GetEffectiveSelectedOptions(element)
                .Select(HtmlFormControlSemantics.GetOptionLabel)
                .Where(value => value.Length > 0)
                .ToArray();
            AddMultilineControlText(visuals, string.Join("\n", values), x, y, width, height, style, false, source, softWrap: false);
            return;
        }

        IElement? selected = HtmlFormControlSemantics.GetEffectiveSelectedOptions(element).SingleOrDefault();
        string value = selected == null ? string.Empty : HtmlFormControlSemantics.GetOptionLabel(selected);
        AddSingleLineControlText(visuals, value, x, y, Math.Max(1D, width - 16D), height, style, false, OfficeTextAlignment.Left, source);

        double arrowWidth = Math.Min(8D, width * 0.12D);
        double arrowHeight = Math.Max(3D, arrowWidth * 0.55D);
        double arrowX = x + width - arrowWidth - 3D;
        double arrowY = y + (height - arrowHeight) / 2D;
        OfficeShape arrow = OfficeShape.Polygon(
            new OfficePoint(0D, 0D),
            new OfficePoint(arrowWidth, 0D),
            new OfficePoint(arrowWidth / 2D, arrowHeight));
        arrow.FillColor = style.Color;
        arrow.StrokeWidth = 0D;
        visuals.Add(new HtmlRenderShape(arrow, arrowX, arrowY, visuals.Count, source: source + ":arrow"));
    }

    private static void AddSingleLineControlText(
        ICollection<HtmlRenderVisual> visuals,
        string text,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBoxStyle style,
        bool placeholder,
        OfficeTextAlignment alignment,
        string source) {
        if (text.Length == 0 || width <= 0D || height <= 0D) return;
        double lineHeight = Math.Min(style.LineHeight, height);
        double textY = y + Math.Max(0D, (height - lineHeight) / 2D);
        visuals.Add(new HtmlRenderText(
            text,
            x,
            textY,
            Math.Max(0.01D, width),
            Math.Max(0.01D, lineHeight),
            style.Font,
            placeholder ? ControlPlaceholderColor : style.Color,
            alignment,
            lineHeight,
            visuals.Count,
            linkUri: null,
            source: source,
            semanticRole: "form-control",
            layoutY: null,
            semanticNodeId: null,
            textAdvanceWidth: null,
            bidiVisualOrderResolved: false,
            semanticFragmentOrder: null,
            logicalTextOrder: null,
            underlineStyle: style.UnderlineStyle,
            strikethroughStyle: style.StrikethroughStyle,
            baseline: style.Baseline,
            decorationColor: style.DecorationColor,
            featureSettings: style.TextFeatureSettings,
            fontPalette: style.FontPalette));
    }

    private void AddMultilineControlText(
        ICollection<HtmlRenderVisual> visuals,
        string text,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBoxStyle style,
        bool placeholder,
        string source,
        bool softWrap) {
        if (text.Length == 0 || width <= 0D || height <= 0D) return;
        IReadOnlyList<string> lines = WrapControlText(text, width, style, softWrap);
        double lineHeight = Math.Max(0.01D, style.LineHeight);
        int maximumLines = Math.Max(1, (int)Math.Floor(height / lineHeight));
        for (int index = 0; index < Math.Min(lines.Count, maximumLines); index++) {
            string line = lines[index];
            if (line.Length == 0) continue;
            visuals.Add(new HtmlRenderText(
                line,
                x,
                y + index * lineHeight,
                Math.Max(0.01D, width),
                Math.Min(lineHeight, Math.Max(0.01D, height - index * lineHeight)),
                style.Font,
                placeholder ? ControlPlaceholderColor : style.Color,
                OfficeTextAlignment.Left,
                lineHeight,
                visuals.Count,
                linkUri: null,
                source: source,
                semanticRole: "form-control",
                layoutY: null,
                semanticNodeId: null,
                textAdvanceWidth: null,
                bidiVisualOrderResolved: false,
                semanticFragmentOrder: null,
                logicalTextOrder: null,
                underlineStyle: style.UnderlineStyle,
                strikethroughStyle: style.StrikethroughStyle,
                baseline: style.Baseline,
                decorationColor: style.DecorationColor,
                featureSettings: style.TextFeatureSettings,
                fontPalette: style.FontPalette));
        }
    }

    private IReadOnlyList<string> WrapControlText(string text, double width, HtmlRenderBoxStyle style, bool softWrap) {
        var result = new List<string>();
        foreach (string logicalLine in text.Split('\n')) {
            if (!softWrap || logicalLine.Length == 0 || MeasureInlineText(logicalLine, style) <= width + 0.0001D) {
                result.Add(logicalLine);
                continue;
            }

            string remaining = logicalLine;
            while (remaining.Length > 0) {
                int fit = 0;
                int lastWhitespace = -1;
                TextElementEnumerator elements = StringInfo.GetTextElementEnumerator(remaining);
                while (elements.MoveNext()) {
                    int end = elements.ElementIndex + elements.GetTextElement().Length;
                    if (MeasureInlineText(remaining.Substring(0, end), style) > width + 0.0001D) break;
                    fit = end;
                    if (char.IsWhiteSpace(remaining[end - 1])) lastWhitespace = end;
                }
                if (fit == 0) {
                    elements = StringInfo.GetTextElementEnumerator(remaining);
                    if (!elements.MoveNext()) break;
                    fit = elements.GetTextElement().Length;
                }
                int take = lastWhitespace > 0 ? lastWhitespace : fit;
                result.Add(remaining.Substring(0, take));
                remaining = remaining.Substring(take);
            }
        }
        return result;
    }

}
