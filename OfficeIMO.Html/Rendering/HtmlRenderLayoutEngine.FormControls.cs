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
        if (style.MinWidth.HasValue) {
            double minimum = style.MinWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
            boxWidth = Math.Max(boxWidth, minimum);
        }
        if (style.MaxWidth.HasValue) {
            double maximum = style.MaxWidth.Value + (style.BorderBox ? 0D : style.HorizontalInsets);
            boxWidth = Math.Min(boxWidth, maximum);
        }
        return Math.Max(1D, Math.Min(availableWidth, boxWidth));
    }

    private static double ResolveFormControlBoxHeight(HtmlRenderBoxStyle style, double defaultContentHeight) {
        double contentHeight = style.ExplicitHeight ?? defaultContentHeight;
        double boxHeight = style.BorderBox && style.ExplicitHeight.HasValue
            ? contentHeight
            : contentHeight + style.VerticalInsets;
        if (style.MinHeight.HasValue) {
            double minimum = style.MinHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets);
            boxHeight = Math.Max(boxHeight, minimum);
        }
        if (style.MaxHeight.HasValue) {
            double maximum = style.MaxHeight.Value + (style.BorderBox ? 0D : style.VerticalInsets);
            boxHeight = Math.Min(boxHeight, maximum);
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
            AddMultilineControlText(visuals, text, contentX, contentY, contentWidth, contentHeight, style, placeholder, source);
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
            AddMultilineControlText(visuals, string.Join("\n", values), x, y, width, height, style, false, source);
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

    private static void AddMultilineControlText(
        ICollection<HtmlRenderVisual> visuals,
        string text,
        double x,
        double y,
        double width,
        double height,
        HtmlRenderBoxStyle style,
        bool placeholder,
        string source) {
        if (text.Length == 0 || width <= 0D || height <= 0D) return;
        string[] lines = text.Split('\n');
        double lineHeight = Math.Max(0.01D, style.LineHeight);
        int maximumLines = Math.Max(1, (int)Math.Floor(height / lineHeight));
        for (int index = 0; index < Math.Min(lines.Length, maximumLines); index++) {
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

        int nodeId = GetSemanticNodeId(element);
        string mappingName = NormalizeControlText(element.GetAttribute("name"));
        if (mappingName.Length == 0) mappingName = NormalizeControlText(element.GetAttribute("id"));
        if (mappingName.Length == 0) mappingName = "html-field-" + nodeId.ToString(CultureInfo.InvariantCulture);
        string name = fieldKind == HtmlRenderFormFieldKind.RadioButton
            ? ResolveRadioFieldName(element, mappingName, nodeId)
            : ResolveUniqueFormFieldName(mappingName, nodeId);

        string value = HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty;
        IReadOnlyList<string> values = Array.Empty<string>();
        IReadOnlyList<string> options = Array.Empty<string>();
        IReadOnlyList<string> optionValues = Array.Empty<string>();
        string? radioOption = null;
        bool selected = fieldKind == HtmlRenderFormFieldKind.CheckBox || fieldKind == HtmlRenderFormFieldKind.RadioButton
            ? HtmlFormControlSemantics.IsEffectivelyChecked(element)
            : false;
        bool multiple = tag == "select" && element.HasAttribute("multiple");

        if (fieldKind == HtmlRenderFormFieldKind.Choice) {
            ResolveChoiceFieldValues(element, out options, out optionValues, out values);
            value = values.FirstOrDefault() ?? string.Empty;
        } else if (fieldKind == HtmlRenderFormFieldKind.RadioButton) {
            radioOption = ResolveRadioOptionToken(element, name, value, nodeId);
            value = radioOption;
        } else if (fieldKind == HtmlRenderFormFieldKind.CheckBox) {
            value = ResolveButtonOptionToken(value, nodeId);
        }

        int? maximumLength = null;
        if (HtmlFormControlSemantics.IsLengthApplicable(tag, type)
            && HtmlFormControlSemantics.TryParseLengthConstraint(element.GetAttribute("maxlength"), out int parsedMaximumLength)
            && parsedMaximumLength > 0) {
            maximumLength = parsedMaximumLength;
        }

        bool disabled = HtmlFormControlSemantics.IsEffectivelyDisabled(element);
        bool readOnly = disabled || element.HasAttribute("readonly") && HtmlFormControlSemantics.IsReadOnlyStateApplicable(tag, type);
        bool required = element.HasAttribute("required") && HtmlFormControlSemantics.IsRequiredStateApplicable(tag, type);
        string alternateName = ResolveFormFieldAccessibleName(element, name);
        OfficeColor? borderColor = style.BorderWidth > 0D && style.BorderStyle != "none" ? style.BorderColor : null;
        formField = new HtmlRenderFormField(
            fieldKind,
            name,
            mappingName,
            value,
            values,
            options,
            optionValues,
            radioOption,
            selected,
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
            style.Alignment,
            style.BackgroundColor,
            borderColor,
            style.BorderWidth,
            x,
            y,
            width,
            height,
            fallbackVisuals,
            paintOrder: 0,
            source);
        return true;
    }

    private void ReportTransformedFormFieldFallback(string source, string detail) {
        if (!_reportedTransformedFormFieldFallbacks.Add(source)) return;
        _diagnostics.Add(
            ComponentName,
            HtmlRenderDiagnosticCodes.FormFieldTransformStaticFallback,
            "A transformed HTML form control was rendered as transformed static content because PDF widget annotations cannot preserve the authored appearance.",
            HtmlDiagnosticSeverity.Warning,
            source,
            detail,
            OfficeConversionLossKind.Approximation);
    }

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

    private static void ResolveChoiceFieldValues(
        IElement select,
        out IReadOnlyList<string> options,
        out IReadOnlyList<string> optionValues,
        out IReadOnlyList<string> values) {
        IElement[] optionElements = select.QuerySelectorAll("option").ToArray();
        var labels = new List<string>(optionElements.Length);
        var exports = new List<string>(optionElements.Length);
        var selectedExports = new List<string>();
        IReadOnlyList<IElement> selectedOptions = HtmlFormControlSemantics.GetEffectiveSelectedOptions(select);
        for (int index = 0; index < optionElements.Length; index++) {
            IElement option = optionElements[index];
            string label = NormalizeControlText(HtmlFormControlSemantics.GetOptionLabel(option));
            if (label.Length == 0) label = NormalizeControlText(option.GetAttribute("value"));
            if (label.Length == 0) label = "Option " + (index + 1).ToString(CultureInfo.InvariantCulture);
            string export = HtmlFormControlSemantics.GetOptionValue(option);
            labels.Add(label);
            exports.Add(export);
            if (selectedOptions.Contains(option) && !selectedExports.Contains(export, StringComparer.Ordinal)) selectedExports.Add(export);
        }
        options = labels;
        optionValues = exports;
        values = selectedExports;
    }

    private string ResolveRadioOptionToken(IElement element, string fieldName, string value, int nodeId) {
        string owner = ResolveFormOwnerKey(element);
        string key = owner + "\n" + fieldName;
        if (!_radioOptionTokens.TryGetValue(key, out HashSet<string>? used)) {
            used = new HashSet<string>(StringComparer.Ordinal);
            _radioOptionTokens[key] = used;
        }
        string candidate = IsPdfNameValue(value) && !string.Equals(value, "Off", StringComparison.Ordinal)
            ? value
            : "option-" + nodeId.ToString(CultureInfo.InvariantCulture);
        if (used.Add(candidate)) return candidate;
        candidate += "-" + nodeId.ToString(CultureInfo.InvariantCulture);
        used.Add(candidate);
        return candidate;
    }

    private string ResolveRadioFieldName(IElement element, string mappingName, int nodeId) {
        string owner = ResolveFormOwnerKey(element);
        string key = owner + "\n" + mappingName;
        if (_radioFieldNames.TryGetValue(key, out string? name)) return name;
        name = ResolveUniqueFormFieldName(mappingName, nodeId);
        _radioFieldNames[key] = name;
        return name;
    }

    private string ResolveFormOwnerKey(IElement element) {
        IElement? owner = HtmlFormControlSemantics.ResolveFormOwner(element);
        if (owner == null) return "none";
        string id = owner.GetAttribute("id") ?? string.Empty;
        return id.Length > 0 ? "id:" + id : "node:" + GetSemanticNodeId(owner).ToString(CultureInfo.InvariantCulture);
    }

    private string ResolveUniqueFormFieldName(string name, int nodeId) {
        if (_formFieldNames.Add(name)) return name;
        string candidate = name + "-" + nodeId.ToString(CultureInfo.InvariantCulture);
        while (!_formFieldNames.Add(candidate)) candidate += "-field";
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
