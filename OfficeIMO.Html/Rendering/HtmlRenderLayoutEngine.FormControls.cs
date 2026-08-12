using System.Globalization;
using AngleSharp.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private static readonly OfficeColor ControlBorderColor = OfficeColor.FromRgb(118, 118, 118);
    private static readonly OfficeColor ControlDisabledBorderColor = OfficeColor.FromRgb(180, 180, 180);
    private static readonly OfficeColor ControlDisabledFillColor = OfficeColor.FromRgb(242, 242, 242);
    private static readonly OfficeColor ControlPlaceholderColor = OfficeColor.FromRgb(105, 105, 105);
    private static readonly OfficeColor ControlAccentColor = OfficeColor.FromRgb(0, 95, 184);

    private static bool IsFormControlElement(string tag) =>
        tag == "input"
        || tag == "select"
        || tag == "textarea"
        || tag == "button"
        || tag == "progress"
        || tag == "meter";

    private double ResolveFormControlOuterWidth(IElement element, HtmlRenderBoxStyle style, double availableWidth) {
        if (IsInputType(element, "image")) {
            return Math.Min(availableWidth, ResolveFloatingImageOuterWidth(element, style));
        }

        HtmlRenderBoxStyle controlStyle = CreateFormControlStyle(element, style);
        double defaultContentWidth = ResolveDefaultFormControlContentWidth(element, controlStyle);
        double availableBoxWidth = Math.Max(1D, availableWidth - controlStyle.MarginLeft - controlStyle.MarginRight);
        double boxWidth = ResolveFormControlBoxWidth(controlStyle, defaultContentWidth, availableBoxWidth);
        return Math.Max(1D, Math.Min(availableWidth, controlStyle.MarginLeft + boxWidth + controlStyle.MarginRight));
    }

    private HtmlRenderFlowBlock LayoutFormControl(IElement element, double containingWidth, HtmlRenderBoxStyle authoredStyle) {
        if (IsInputType(element, "image")) {
            return LayoutImage(element, containingWidth, authoredStyle);
        }

        string source = HtmlRenderStyleResolver.DescribeSource(element);
        HtmlRenderBoxStyle style = CreateFormControlStyle(element, authoredStyle);
        double availableWidth = Math.Max(1D, containingWidth - style.MarginLeft - style.MarginRight);
        double defaultContentWidth = ResolveDefaultFormControlContentWidth(element, style);
        double boxWidth = ResolveFormControlBoxWidth(style, defaultContentWidth, availableWidth);
        double defaultContentHeight = ResolveDefaultFormControlContentHeight(element, style);
        double boxHeight = ResolveFormControlBoxHeight(style, defaultContentHeight);
        double x = style.MarginLeft;
        double y = style.MarginTop;

        var visuals = new List<HtmlRenderVisual>();
        AddBoxPaint(visuals, style, x, y, boxWidth, boxHeight, element);
        if (style.PaintVisible) {
            AddFormControlContent(visuals, element, style, x, y, boxWidth, boxHeight, source);
            AddBoxOutlinePaint(visuals, style, x, y, boxWidth, boxHeight, element);
            if (TryCreateFormFieldVisual(element, style, x, y, boxWidth, boxHeight, visuals, source, out HtmlRenderFormField? formField)) {
                visuals = new List<HtmlRenderVisual> { formField! };
            }
        } else {
            visuals.Clear();
        }

        double height = style.MarginTop + boxHeight + style.MarginBottom;
        return new HtmlRenderFlowBlock(
            containingWidth,
            Math.Max(0.01D, height),
            visuals,
            style.BreakBefore,
            style.BreakAfter,
            avoidBreakInside: true,
            source,
            pageName: style.PageName);
    }

    private HtmlRenderBoxStyle CreateFormControlStyle(IElement element, HtmlRenderBoxStyle authoredStyle) {
        HtmlRenderBoxStyle style = authoredStyle.Clone();
        bool compact = IsCompactChoiceControl(element);
        bool range = IsInputType(element, "range");

        if (!style.BorderDeclared && !range) {
            style.Borders = HtmlRenderBorderEdges.Uniform(
                1D,
                "solid",
                element.HasAttribute("disabled") ? ControlDisabledBorderColor : ControlBorderColor);
            style.BorderDeclared = true;
        }
        if (style.BackgroundColor == null) {
            style.BackgroundColor = element.HasAttribute("disabled")
                ? ControlDisabledFillColor
                : OfficeColor.White;
        }
        if (!compact && !range && style.PaddingLeft == 0D && style.PaddingRight == 0D) {
            style.PaddingLeft = 6D;
            style.PaddingRight = 6D;
        }
        if (!compact && !range && style.PaddingTop == 0D && style.PaddingBottom == 0D) {
            style.PaddingTop = 4D;
            style.PaddingBottom = 4D;
        }
        if (!compact && style.BorderRadius == "0") style.BorderRadius = "3px";
        style.AvoidBreakInside = true;
        style.SemanticRole = "form-control";
        return style;
    }

    private double ResolveDefaultFormControlContentWidth(IElement element, HtmlRenderBoxStyle style) {
        string tag = element.TagName.ToLowerInvariant();
        string type = NormalizeInputType(element);
        if (tag == "input" && (type == "checkbox" || type == "radio")) return 14D;
        if (tag == "input" && type == "color") return 32D;
        if (tag == "input" && type == "range") return 144D;
        if (tag == "progress" || tag == "meter") return 144D;
        if (tag == "textarea") {
            int columns = ParsePositiveInteger(element.GetAttribute("cols"), 20, 1, 200);
            return Math.Max(80D, MeasureInlineText(new string('0', columns), style));
        }
        if (tag == "button" || tag == "input" && IsButtonInputType(type)) {
            string label = ResolveButtonLabel(element, type);
            return Math.Max(44D, MeasureInlineText(label, style) + 12D);
        }
        if (tag == "select") {
            string longest = element.QuerySelectorAll("option")
                .Select(HtmlFormControlSemantics.GetOptionLabel)
                .OrderByDescending(text => text.Length)
                .FirstOrDefault() ?? string.Empty;
            return Math.Max(108D, MeasureInlineText(longest, style) + 24D);
        }
        if (tag == "input" && type == "file") return 220D;
        return 168D;
    }

    private static double ResolveDefaultFormControlContentHeight(IElement element, HtmlRenderBoxStyle style) {
        string tag = element.TagName.ToLowerInvariant();
        string type = NormalizeInputType(element);
        if (tag == "input" && (type == "checkbox" || type == "radio")) return 14D;
        if (tag == "input" && type == "color") return 22D;
        if (tag == "input" && type == "range" || tag == "progress" || tag == "meter") return 14D;
        if (tag == "textarea") {
            int rows = ParsePositiveInteger(element.GetAttribute("rows"), 2, 1, 100);
            return Math.Max(style.LineHeight, rows * style.LineHeight);
        }
        int selectDisplaySize = tag == "select" ? HtmlFormControlSemantics.GetSelectDisplaySize(element) : 1;
        if (tag == "select" && (element.HasAttribute("multiple") || selectDisplaySize > 1)) {
            int rows = Math.Max(2, Math.Min(20, selectDisplaySize));
            return Math.Max(style.LineHeight, rows * style.LineHeight);
        }
        return Math.Max(style.LineHeight, 20D);
    }

    private static double ResolveFormControlBoxWidth(HtmlRenderBoxStyle style, double defaultContentWidth, double availableWidth) {
        double contentWidth = style.ExplicitWidth ?? defaultContentWidth;
        double boxWidth = style.BorderBox && style.ExplicitWidth.HasValue
            ? contentWidth
            : contentWidth + style.HorizontalInsets;
        if (style.MaxWidth.HasValue) {
            double maximum = style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
            boxWidth = Math.Min(boxWidth, maximum);
        }
        if (style.MinWidth.HasValue) {
            double minimum = style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
            boxWidth = Math.Max(boxWidth, minimum);
        }
        return Math.Max(1D, Math.Min(availableWidth, boxWidth));
    }

    private static double ResolveFormControlBoxHeight(HtmlRenderBoxStyle style, double defaultContentHeight) {
        double contentHeight = style.ExplicitHeight ?? defaultContentHeight;
        double boxHeight = style.BorderBox && style.ExplicitHeight.HasValue
            ? contentHeight
            : contentHeight + style.VerticalInsets;
        if (style.MaxHeight.HasValue) {
            double maximum = style.MaxHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets);
            boxHeight = Math.Min(boxHeight, maximum);
        }
        if (style.MinHeight.HasValue) {
            double minimum = style.MinHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets);
            boxHeight = Math.Max(boxHeight, minimum);
        }
        return Math.Max(1D, boxHeight);
    }

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
            if (HtmlFormControlSemantics.IsEffectivelyChecked(element)) AddCheckboxMark(visuals, contentX, contentY, contentWidth, contentHeight, source);
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
            string text = NormalizeControlMultilineText(element.TextContent);
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
                ? NormalizeControlText(HtmlFormControlSemantics.GetValues(element).FirstOrDefault())
                : NormalizeControlText(element.GetAttribute("value"));
            if (type == "password" && value.Length > 0) value = new string('*', Math.Min(32, value.Length));
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
            source: source,
            semanticRole: "form-control"));
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
                source: source,
                semanticRole: "form-control"));
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

    private bool TryCreateFormFieldVisual(
        IElement element,
        HtmlRenderBoxStyle style,
        double x,
        double y,
        double width,
        double height,
        IReadOnlyList<HtmlRenderVisual> fallbackVisuals,
        string source,
        out HtmlRenderFormField? formField) {
        formField = null;
        if (!string.Equals(style.Transform, "none", StringComparison.OrdinalIgnoreCase)) {
            ReportTransformedFormFieldFallback(source, "transform=" + style.Transform);
            return false;
        }
        string tag = element.LocalName.ToLowerInvariant();
        string type = tag == "input" ? NormalizeInputType(element) : tag;
        HtmlRenderFormFieldKind fieldKind;
        if (tag == "textarea") fieldKind = HtmlRenderFormFieldKind.Text;
        else if (tag == "select") fieldKind = HtmlRenderFormFieldKind.Choice;
        else if (tag == "input" && type == "checkbox") fieldKind = HtmlRenderFormFieldKind.CheckBox;
        else if (tag == "input" && type == "radio") fieldKind = HtmlRenderFormFieldKind.RadioButton;
        else if (tag == "input" && IsInteractiveTextInputType(type)) fieldKind = HtmlRenderFormFieldKind.Text;
        else return false;
        if (tag == "input" && type == "file" && element.HasAttribute("multiple")) {
            ReportMultipleFileSelectionFallback(source);
            return false;
        }

        string? authoredName = element.GetAttribute("name");
        if (authoredName != null && authoredName.Length > 0 && string.IsNullOrWhiteSpace(authoredName)) {
            ReportBlankFormFieldNameFallback(source);
            return false;
        }
        if (_staticRepeatedControlGroupKeys.TryGetValue(element, out string? staticGroupKey)) {
            ReportRepeatedFormControlNameFallback(source, staticGroupKey);
            return false;
        }
        if (fieldKind == HtmlRenderFormFieldKind.RadioButton && _blankValueRadioGroupKeys.TryGetValue(element, out staticGroupKey)) {
            ReportBlankButtonValueFallback(source, staticGroupKey);
            return false;
        }
        if (fieldKind == HtmlRenderFormFieldKind.RadioButton && _staticRadioGroupKeys.TryGetValue(element, out staticGroupKey)) {
            ReportDuplicateRadioValueFallback(source, staticGroupKey);
            return false;
        }
        if (fieldKind == HtmlRenderFormFieldKind.RadioButton && _mixedDisabledRadioGroupKeys.TryGetValue(element, out staticGroupKey)) {
            ReportMixedDisabledRadioGroupFallback(source, staticGroupKey);
            return false;
        }
        if (fieldKind == HtmlRenderFormFieldKind.RadioButton && _transparentRadioGroupKeys.TryGetValue(element, out staticGroupKey)) {
            ReportTransparentFormFieldPaintFallback(source, staticGroupKey);
            return false;
        }
        if (fieldKind == HtmlRenderFormFieldKind.RadioButton && _backgroundImageRadioGroupKeys.TryGetValue(element, out staticGroupKey)) {
            ReportBackgroundImageFormFieldFallback(source, staticGroupKey);
            return false;
        }

        int nodeId = GetSemanticNodeId(element);
        string mappingName = authoredName is { Length: > 0 } ? authoredName : string.Empty;
        string fieldName = mappingName;
        if (fieldName.Length == 0) fieldName = NormalizeControlText(element.GetAttribute("id"));
        if (fieldName.Length == 0) fieldName = "html-field-" + nodeId.ToString(CultureInfo.InvariantCulture);
        string name = fieldKind == HtmlRenderFormFieldKind.RadioButton
            ? ResolveRadioFieldName(element, fieldName, nodeId)
            : ResolveUniqueFormFieldName(element, fieldName, nodeId);

        string value = HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty;
        string placeholder = value.Length == 0 && HtmlFormControlSemantics.IsPlaceholderApplicable(tag, type)
            ? tag == "textarea"
                ? NormalizeControlMultilineText(element.GetAttribute("placeholder") ?? string.Empty)
                : NormalizeControlText(element.GetAttribute("placeholder"))
            : string.Empty;
        IReadOnlyList<string> values = Array.Empty<string>();
        IReadOnlyList<string> options = Array.Empty<string>();
        IReadOnlyList<string> optionValues = Array.Empty<string>();
        IReadOnlyList<int> selectedOptionIndices = Array.Empty<int>();
        string? radioOption = null;
        bool selected = fieldKind == HtmlRenderFormFieldKind.CheckBox || fieldKind == HtmlRenderFormFieldKind.RadioButton
            ? HtmlFormControlSemantics.IsEffectivelyChecked(element)
            : false;
        bool multiple = tag == "select" && element.HasAttribute("multiple");

        if (fieldKind == HtmlRenderFormFieldKind.Choice) {
            if (element.QuerySelectorAll("option").Any(HtmlFormControlSemantics.IsOptionEffectivelyDisabled)) {
                ReportDisabledChoiceOptionFallback(source);
                return false;
            }
            ResolveChoiceFieldValues(element, out options, out optionValues, out values, out selectedOptionIndices, out bool hasDuplicateSelectedValues, out bool hasAmbiguousSelectedValue);
            if (options.Count == 0) {
                ReportEmptyChoiceOptionsFallback(source);
                return false;
            }
            if (options.Any(label => label.Length == 0)) {
                ReportBlankChoiceLabelFallback(source);
                return false;
            }
            if (multiple && hasDuplicateSelectedValues) {
                ReportDuplicateSelectedChoiceValueFallback(source);
                return false;
            }
            if (!multiple && HtmlFormControlSemantics.GetSelectDisplaySize(element) == 1 && hasAmbiguousSelectedValue) {
                ReportDuplicateSelectedChoiceValueFallback(source);
                return false;
            }
            value = values.FirstOrDefault() ?? string.Empty;
        } else if (fieldKind == HtmlRenderFormFieldKind.RadioButton) {
            if (string.IsNullOrWhiteSpace(value)) {
                ReportBlankButtonValueFallback(source, null);
                return false;
            }
            radioOption = ResolveRadioOptionToken(value, nodeId);
        } else if (fieldKind == HtmlRenderFormFieldKind.CheckBox) {
            if (string.IsNullOrWhiteSpace(value)) {
                ReportBlankButtonValueFallback(source, null);
                return false;
            }
            radioOption = ResolveButtonOptionToken(value, nodeId);
        }

        int? maximumLength = null;
        if (HtmlFormControlSemantics.IsLengthApplicable(tag, type)
            && HtmlFormControlSemantics.TryParseLengthConstraint(element.GetAttribute("maxlength"), out int parsedMaximumLength)) {
            if (parsedMaximumLength == 0) {
                ReportZeroMaximumLengthFallback(source);
                return false;
            }
            maximumLength = parsedMaximumLength;
        }

        bool disabled = HtmlFormControlSemantics.IsEffectivelyDisabled(element);
        bool readOnly = disabled || element.HasAttribute("readonly") && HtmlFormControlSemantics.IsReadOnlyStateApplicable(tag, type);
        bool required = element.HasAttribute("required")
            && HtmlFormControlSemantics.IsRequiredStateApplicable(tag, type)
            && !disabled
            && !readOnly;
        string alternateName = ResolveFormFieldAccessibleName(element, name);
        OfficeColor? borderColor = style.BorderWidth > 0D && style.BorderStyle != "none" ? style.BorderColor : null;
        if (borderColor.HasValue && style.BorderStyle != "solid" && style.BorderStyle != "dashed") {
            ReportUnsupportedFormFieldBorderStyleFallback(source, style.BorderStyle);
            return false;
        }
        if (HasUnsupportedInteractiveFieldTransparency(style.Color, style.BackgroundColor, borderColor)) {
            ReportTransparentFormFieldPaintFallback(source, null);
            return false;
        }
        if (style.BackgroundImageLayers.Count > 0) {
            ReportBackgroundImageFormFieldFallback(source, null);
            return false;
        }
        HtmlResolvedBorderRadii resolvedRadii = ResolveBoxRadii(style, width, height, element, source);
        if (!resolvedRadii.IsZero && !resolvedRadii.IsUniformCircular) {
            ReportNonUniformFormFieldRadiusFallback(source);
            return false;
        }
        formField = new HtmlRenderFormField(
            fieldKind,
            name,
            mappingName,
            value,
            placeholder,
            values,
            options,
            optionValues,
            selectedOptionIndices,
            radioOption,
            selected,
            disabled,
            readOnly,
            required,
            tag == "textarea",
            tag == "input" && type == "password",
            tag == "input" && type == "file",
            tag == "select" && !multiple && HtmlFormControlSemantics.GetSelectDisplaySize(element) == 1,
            multiple,
            maximumLength,
            alternateName,
            style.Font,
            style.Color,
            ControlPlaceholderColor,
            style.Alignment,
            style.BackgroundColor,
            borderColor,
            style.BorderStyle,
            style.BorderWidth,
            resolvedRadii.UniformRadius,
            x,
            y,
            width,
            height,
            fallbackVisuals,
            paintOrder: 0,
            source);
        return true;
    }

    private void ReportUnsupportedFormFieldBorderStyleFallback(string source, string borderStyle) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldBorderStyleStaticFallback,
            "An HTML form control used faithful static rendering because its authored border style cannot be represented by a PDF widget appearance.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "border-style=" + borderStyle,
            OfficeConversionLossKind.Approximation);
    }

    private void ReportTransformedFormFieldFallback(string source, string detail) {
        if (!_reportedTransformedFormFieldFallbacks.Add(source)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldTransformStaticFallback,
            "An HTML form control inside a transformed or translucent paint group was rendered as static content because PDF widget annotations cannot preserve the authored appearance.",
            HtmlDiagnosticSeverity.Warning,
            source,
            detail,
            OfficeConversionLossKind.Approximation);
    }

    private void ReportDuplicateRadioValueFallback(string source, string groupKey) {
        if (!_reportedStaticRadioGroups.Add(groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.RadioDuplicateValueStaticFallback,
            "An HTML radio group with duplicate submitted values was rendered as static content because PDF radio appearance-state values must be unique.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportMixedDisabledRadioGroupFallback(string source, string groupKey) {
        if (!_reportedStaticRadioGroups.Add(groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.RadioMixedDisabledStateStaticFallback,
            "An HTML radio group mixing enabled and disabled options was rendered as static content because PDF radio widgets cannot preserve disabled state per option.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportZeroMaximumLengthFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldZeroMaximumLengthStaticFallback,
            "An HTML text control with maxlength=0 was rendered as static content because PDF /MaxLen must be positive.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "maxlength=0",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportMultipleFileSelectionFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FileMultipleSelectionStaticFallback,
            "An HTML multiple-file input was rendered as static content because PDF file-select fields cannot preserve multiple-file selection semantics.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "input[type=file][multiple]",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportBlankFormFieldNameFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldBlankNameStaticFallback,
            "An HTML form control with a whitespace-only name was rendered as static content because PDF form field names must contain a non-whitespace character.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "whitespace-only name",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportBlankButtonValueFallback(string source, string? groupKey) {
        if (groupKey != null && !_reportedStaticRadioGroups.Add(groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldBlankButtonValueStaticFallback,
            "An HTML checkbox or radio control with an empty or whitespace-only submitted value was rendered as static content because PDF button export values must contain a non-whitespace character.",
            HtmlDiagnosticSeverity.Warning,
            source,
            groupKey == null ? "blank submitted value" : "group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportDuplicateSelectedChoiceValueFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ChoiceDuplicateSelectedValueStaticFallback,
            "An HTML multi-select with duplicate selected submitted values was rendered as static content because a value-only PDF choice selection cannot preserve both selected option identities.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "duplicate selected option values",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportDisabledChoiceOptionFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ChoiceDisabledOptionStaticFallback,
            "An HTML select containing disabled options was rendered as static content because PDF choice fields cannot preserve disabled state per option.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "disabled option",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportRepeatedFormControlNameFallback(string source, string groupKey) {
        if (!_reportedStaticRepeatedControlGroups.Add(groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldRepeatedNameStaticFallback,
            "HTML controls sharing one submitted name were rendered as static content because distinct PDF fields cannot preserve repeated form-data entries under one field name.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "name=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportBlankChoiceLabelFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ChoiceBlankLabelStaticFallback,
            "An HTML select containing a blank option label was rendered as static content because PDF choice fields require non-empty display labels.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "blank option label",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportEmptyChoiceOptionsFallback(string source) {
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.ChoiceEmptyOptionsStaticFallback,
            "An empty HTML select was rendered as static content because interactive PDF choice fields require at least one option.",
            HtmlDiagnosticSeverity.Warning,
            source,
            "empty option list",
            OfficeConversionLossKind.Approximation);
    }

    private void ReportTransparentFormFieldPaintFallback(string source, string? groupKey) {
        if (groupKey != null && !_reportedStaticRadioGroups.Add("transparent\n" + groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldColorTransparencyStaticFallback,
            "An HTML form control with translucent paint was rendered as static content because generated PDF widget appearances cannot preserve color alpha.",
            HtmlDiagnosticSeverity.Warning,
            source,
            groupKey == null
                ? "translucent form-control paint"
                : "translucent radio-group paint; group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private void ReportBackgroundImageFormFieldFallback(string source, string? groupKey) {
        if (groupKey != null && !_reportedStaticRadioGroups.Add("background-image\n" + groupKey)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldBackgroundImageStaticFallback,
            "An HTML form control with background-image paint was rendered as static content because generated PDF widget appearances cannot preserve background layers.",
            HtmlDiagnosticSeverity.Warning,
            source,
            groupKey == null
                ? "form-control background-image"
                : "radio-group background-image; group=" + groupKey.Substring(groupKey.LastIndexOf('\n') + 1),
            OfficeConversionLossKind.Approximation);
    }

    private static bool HasUnsupportedInteractiveFieldTransparency(
        OfficeColor textColor,
        OfficeColor? backgroundColor,
        OfficeColor? borderColor) =>
        textColor.A < byte.MaxValue
        || HasPartialAlpha(backgroundColor)
        || HasPartialAlpha(borderColor);

    private static bool HasPartialAlpha(OfficeColor? color) =>
        color.HasValue && color.Value.A > 0 && color.Value.A < byte.MaxValue;

    private static bool IsInteractiveTextInputType(string type) {
        switch (type) {
            case "date":
            case "datetime-local":
            case "email":
            case "file":
            case "month":
            case "number":
            case "password":
            case "search":
            case "tel":
            case "text":
            case "time":
            case "url":
            case "week":
                return true;
            default:
                return false;
        }
    }

    private static string ResolveRadioOptionToken(string value, int nodeId) {
        return IsPdfNameValue(value) && !string.Equals(value, "Off", StringComparison.Ordinal)
            ? value
            : "option-" + nodeId.ToString(CultureInfo.InvariantCulture);
    }

    private void IdentifyStaticRadioGroups() {
        _staticRadioGroupKeys.Clear();
        _blankValueRadioGroupKeys.Clear();
        _mixedDisabledRadioGroupKeys.Clear();
        _transparentRadioGroupKeys.Clear();
        _backgroundImageRadioGroupKeys.Clear();
        _staticRepeatedControlGroupKeys.Clear();
        IElement[] radios = _document.QuerySelectorAll("input")
            .Where(element => NormalizeInputType(element) == "radio")
            .Where(element => element.GetAttribute("name") is { Length: > 0 })
            .ToArray();
        foreach (IGrouping<IElement?, IElement> ownerGroup in radios.GroupBy(HtmlFormControlSemantics.ResolveFormOwner)) {
            foreach (IGrouping<string, IElement> group in ownerGroup.GroupBy(
                         element => element.GetAttribute("name")!,
                         StringComparer.Ordinal)) {
                IElement[] members = group.ToArray();
                bool hasBlankValue = members.Any(element =>
                    string.IsNullOrWhiteSpace(HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty));
                bool hasDuplicateValue = members
                .GroupBy(element => HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty, StringComparer.Ordinal)
                .Any(values => values.Skip(1).Any());
                bool hasDisabled = members.Any(HtmlFormControlSemantics.IsEffectivelyDisabled);
                bool hasEnabled = members.Any(element => !HtmlFormControlSemantics.IsEffectivelyDisabled(element));
                bool hasTransparentPaint = false;
                bool hasBackgroundImage = false;
                foreach (IElement element in members) {
                    HtmlRenderBoxStyle authoredStyle = _styleResolver.Resolve(element, _options.PageWidth);
                    HtmlRenderBoxStyle controlStyle = CreateFormControlStyle(element, authoredStyle);
                    OfficeColor? borderColor = controlStyle.BorderWidth > 0D && controlStyle.BorderStyle != "none"
                        ? controlStyle.BorderColor
                        : null;
                    hasTransparentPaint |= HasUnsupportedInteractiveFieldTransparency(controlStyle.Color, controlStyle.BackgroundColor, borderColor);
                    hasBackgroundImage |= controlStyle.BackgroundImageLayers.Count > 0;
                }
                if (!hasBlankValue && !hasDuplicateValue && !(hasDisabled && hasEnabled) && !hasTransparentPaint && !hasBackgroundImage) continue;
                string key = HtmlRenderStyleResolver.DescribeSource(members[0]) + "\n" + group.Key;
                foreach (IElement element in members) {
                    if (hasBlankValue) _blankValueRadioGroupKeys[element] = key;
                    if (hasDuplicateValue) _staticRadioGroupKeys[element] = key;
                    if (hasDisabled && hasEnabled) _mixedDisabledRadioGroupKeys[element] = key;
                    if (hasTransparentPaint) _transparentRadioGroupKeys[element] = key;
                    if (hasBackgroundImage) _backgroundImageRadioGroupKeys[element] = key;
                }
            }
        }

        IElement[] controls = _document.QuerySelectorAll("input,textarea,select")
            .Where(IsSupportedInteractiveFormControl)
            .Where(element => element.GetAttribute("name") is { Length: > 0 })
            .ToArray();
        foreach (IGrouping<IElement?, IElement> ownerGroup in controls.GroupBy(HtmlFormControlSemantics.ResolveFormOwner)) {
            foreach (IGrouping<string, IElement> group in ownerGroup.GroupBy(
                         element => element.GetAttribute("name")!,
                         StringComparer.Ordinal)) {
                IElement[] members = group.ToArray();
                if (members.Length < 2 || members.All(element => element.LocalName == "input" && NormalizeInputType(element) == "radio")) continue;
                string key = HtmlRenderStyleResolver.DescribeSource(members[0]) + "\n" + group.Key;
                foreach (IElement element in members) _staticRepeatedControlGroupKeys[element] = key;
            }
        }
    }

    private static bool IsSupportedInteractiveFormControl(IElement element) {
        if (element.LocalName == "textarea" || element.LocalName == "select") return true;
        if (element.LocalName != "input") return false;
        string type = NormalizeInputType(element);
        return type == "checkbox" || type == "radio" || IsInteractiveTextInputType(type);
    }

    private string ResolveRadioFieldName(IElement element, string mappingName, int nodeId) {
        string owner = ResolveFormOwnerKey(element);
        string key = owner + "\n" + mappingName;
        if (_radioFieldNames.TryGetValue(key, out string? name)) return name;
        name = ResolveUniqueFormFieldName(element, mappingName, nodeId);
        _radioFieldNames[key] = name;
        return name;
    }

    private string ResolveFormOwnerKey(IElement element) {
        IElement? owner = HtmlFormControlSemantics.ResolveFormOwner(element);
        if (owner == null) return "none";
        string id = owner.GetAttribute("id") ?? string.Empty;
        return id.Length > 0 ? "id:" + id : "node:" + GetSemanticNodeId(owner).ToString(CultureInfo.InvariantCulture);
    }

    private string ResolveUniqueFormFieldName(IElement element, string name, int nodeId) {
        if (_formFieldNamesByElement.TryGetValue(element, out string? existing)) return existing;
        if (_formFieldNames.Add(name)) {
            _formFieldNamesByElement[element] = name;
            return name;
        }
        string candidate = name + "-" + nodeId.ToString(CultureInfo.InvariantCulture);
        while (!_formFieldNames.Add(candidate)) candidate += "-field";
        _formFieldNamesByElement[element] = candidate;
        return candidate;
    }

    private static string ResolveButtonOptionToken(string value, int nodeId) =>
        IsPdfNameValue(value) && !string.Equals(value, "Off", StringComparison.Ordinal)
            ? value
            : "value-" + nodeId.ToString(CultureInfo.InvariantCulture);

    private static bool IsPdfNameValue(string value) {
        if (string.IsNullOrWhiteSpace(value)) return false;
        for (int index = 0; index < value.Length; index++) {
            if (value[index] > 0x7E) return false;
        }
        return true;
    }

    private string ResolveFormFieldAccessibleName(IElement element, string fallbackName) {
        string accessibleName = HtmlAccessibilitySemantics.GetAccessibleName(element);
        if (accessibleName.Length > 0) return accessibleName;
        for (IElement? ancestor = element.ParentElement; ancestor != null; ancestor = ancestor.ParentElement) {
            if (!string.Equals(ancestor.LocalName, "label", StringComparison.OrdinalIgnoreCase)) continue;
            accessibleName = HtmlAccessibilitySemantics.GetAccessibleName(ancestor, includeTextFallback: true);
            if (accessibleName.Length > 0) return accessibleName;
            break;
        }
        string id = element.GetAttribute("id") ?? string.Empty;
        if (id.Length > 0) {
            foreach (IElement label in _document.QuerySelectorAll("label")) {
                if (!string.Equals(label.GetAttribute("for"), id, StringComparison.Ordinal)) continue;
                accessibleName = HtmlAccessibilitySemantics.GetAccessibleName(label, includeTextFallback: true);
                if (accessibleName.Length > 0) return accessibleName;
            }
        }
        string placeholder = NormalizeControlText(element.GetAttribute("placeholder"));
        return placeholder.Length > 0 ? placeholder : fallbackName;
    }

    private static double ResolveNumericFraction(IElement element, double defaultMinimum, double defaultMaximum, double defaultValue) {
        double minimum = ParseFiniteDouble(element.GetAttribute("min"), defaultMinimum);
        double maximum = ParseFiniteDouble(element.GetAttribute("max"), defaultMaximum);
        if (maximum <= minimum) maximum = minimum + 1D;
        double value = ParseFiniteDouble(element.GetAttribute("value"), defaultValue);
        return Math.Max(0D, Math.Min(1D, (value - minimum) / (maximum - minimum)));
    }

    private static string ResolveButtonLabel(IElement element, string type) {
        if (string.Equals(element.TagName, "button", StringComparison.OrdinalIgnoreCase)) {
            string content = NormalizeControlText(element.TextContent);
            return content.Length == 0 ? "Button" : content;
        }
        string value = NormalizeControlText(element.GetAttribute("value"));
        if (value.Length > 0) return value;
        if (type == "submit") return "Submit";
        if (type == "reset") return "Reset";
        return "Button";
    }

    private static bool IsButtonInputType(string type) =>
        type == "button" || type == "submit" || type == "reset";

    private static bool IsCompactChoiceControl(IElement element) {
        if (!string.Equals(element.TagName, "input", StringComparison.OrdinalIgnoreCase)) return false;
        string type = NormalizeInputType(element);
        return type == "checkbox" || type == "radio";
    }

    private static bool IsInputType(IElement element, string type) =>
        string.Equals(element.TagName, "input", StringComparison.OrdinalIgnoreCase)
        && string.Equals(NormalizeInputType(element), type, StringComparison.Ordinal);

    private static string NormalizeInputType(IElement element) {
        if (!string.Equals(element.TagName, "input", StringComparison.OrdinalIgnoreCase)) return string.Empty;
        return HtmlFormControlSemantics.GetEffectiveType("input", element.GetAttribute("type"));
    }

    private static string NormalizeControlText(string? value) =>
        string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static string NormalizeControlMultilineText(string value) =>
        value.Replace("\r\n", "\n").Replace('\r', '\n').Trim();

    private static int ParsePositiveInteger(string? value, int fallback, int minimum, int maximum) =>
        HtmlIntegerSemantics.TryParsePositiveInteger(value, out int parsed)
            ? Math.Max(minimum, Math.Min(maximum, parsed))
            : fallback;

    private static double ParseFiniteDouble(string? value, double fallback) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
        && !double.IsNaN(parsed)
        && !double.IsInfinity(parsed)
            ? parsed
            : fallback;
}
