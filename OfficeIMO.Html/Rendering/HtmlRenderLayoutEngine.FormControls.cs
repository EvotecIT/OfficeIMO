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
            return Math.Max(80D, MeasureText(new string('0', columns), style.Font));
        }
        if (tag == "button" || tag == "input" && IsButtonInputType(type)) {
            string label = ResolveButtonLabel(element, type);
            return Math.Max(44D, MeasureText(label, style.Font) + 12D);
        }
        if (tag == "select") {
            string longest = element.QuerySelectorAll("option")
                .Select(option => NormalizeControlText(option.TextContent))
                .OrderByDescending(text => text.Length)
                .FirstOrDefault() ?? string.Empty;
            return Math.Max(108D, MeasureText(longest, style.Font) + 24D);
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
        if (tag == "select" && (element.HasAttribute("multiple") || ParsePositiveInteger(element.GetAttribute("size"), 1, 1, 100) > 1)) {
            int rows = ParsePositiveInteger(element.GetAttribute("size"), 4, 2, 20);
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
            if (element.HasAttribute("checked")) AddCheckboxMark(visuals, contentX, contentY, contentWidth, contentHeight, source);
            return;
        }
        if (tag == "input" && type == "radio") {
            ReplaceControlBackgroundWithRadio(visuals, boxX, boxY, boxWidth, boxHeight, style, source);
            if (element.HasAttribute("checked")) AddRadioMark(visuals, contentX, contentY, contentWidth, contentHeight, source);
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
            value = NormalizeControlText(element.GetAttribute("value"));
            value = value.Length == 0 ? "Choose file" : value;
        } else {
            value = NormalizeControlText(element.GetAttribute("value"));
            if (type == "password" && value.Length > 0) value = new string('*', Math.Min(32, value.Length));
            if (value.Length == 0) {
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
        double fraction = ResolveNumericFraction(element, 0D, 100D, 50D);
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
        OfficeColor color = HtmlRenderCssValues.TryColor(element.GetAttribute("value") ?? string.Empty, out OfficeColor parsed)
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
        IElement[] options = element.QuerySelectorAll("option").ToArray();
        bool multiple = element.HasAttribute("multiple") || ParsePositiveInteger(element.GetAttribute("size"), 1, 1, 100) > 1;
        if (multiple) {
            string[] values = options
                .Where(option => option.HasAttribute("selected"))
                .DefaultIfEmpty(options.FirstOrDefault()!)
                .Where(option => option != null)
                .Select(option => NormalizeControlText(option.TextContent))
                .Where(value => value.Length > 0)
                .ToArray();
            AddMultilineControlText(visuals, string.Join("\n", values), x, y, width, height, style, false, source);
            return;
        }

        IElement? selected = options.FirstOrDefault(option => option.HasAttribute("selected")) ?? options.FirstOrDefault();
        string value = selected == null ? string.Empty : NormalizeControlText(selected.TextContent);
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
        string type = (element.GetAttribute("type") ?? string.Empty).Trim().ToLowerInvariant();
        return type.Length == 0 ? "text" : type;
    }

    private static string NormalizeControlText(string? value) =>
        string.Join(" ", (value ?? string.Empty).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries));

    private static string NormalizeControlMultilineText(string value) =>
        value.Replace("\r\n", "\n").Replace('\r', '\n').Trim();

    private static int ParsePositiveInteger(string? value, int fallback, int minimum, int maximum) =>
        int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
            ? Math.Max(minimum, Math.Min(maximum, parsed))
            : fallback;

    private static double ParseFiniteDouble(string? value, double fallback) =>
        double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsed)
        && !double.IsNaN(parsed)
        && !double.IsInfinity(parsed)
            ? parsed
            : fallback;
}
