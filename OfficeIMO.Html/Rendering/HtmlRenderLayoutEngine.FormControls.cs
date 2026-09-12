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
        if (tag == "input" && IsInputSizeApplicable(type)) {
            int size = ParsePositiveInteger(element.GetAttribute("size"), 20, 1, int.MaxValue);
            return Math.Max(1D, MeasureInlineText("0", style) * size);
        }
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
        if (HtmlFormControlSemantics.IsIndeterminate(element)) {
            _diagnostics.Add(ComponentName, HtmlRenderDiagnosticCodes.FormFieldIndeterminateStaticFallback,
                "An indeterminate checkbox used static rendering because a PDF checkbox widget cannot preserve the mixed appearance.",
                HtmlDiagnosticSeverity.Warning, source, "input[type=checkbox]", OfficeConversionLossKind.Approximation);
            return false;
        }
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
        if (tag == "textarea" && string.Equals(element.GetAttribute("wrap")?.Trim(), "off", StringComparison.OrdinalIgnoreCase)) {
            ReportNoWrapFormFieldFallback(source);
            return false;
        }
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
        string partialFieldName = NormalizePdfPartialFieldName(fieldName);
        string name = fieldKind == HtmlRenderFormFieldKind.RadioButton
            ? ResolveRadioFieldName(element, fieldName, partialFieldName, nodeId)
            : ResolveUniqueFormFieldName(element, partialFieldName, nodeId);

        string value = HtmlFormControlSemantics.GetValues(element).FirstOrDefault() ?? string.Empty;
        if (tag == "textarea") value = NormalizeControlMultilineText(value);
        bool emptyFileSelect = value.Length == 0 && tag == "input" && type == "file";
        string placeholder = emptyFileSelect
            ? "Choose file"
            : value.Length == 0 && HtmlFormControlSemantics.IsPlaceholderApplicable(tag, type)
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
            if (value.Length > parsedMaximumLength) {
                ReportInitialValueExceedsMaximumLengthFallback(source, parsedMaximumLength, value.Length);
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
        string alternateName = ResolveFormFieldAccessibleName(element, mappingName.Length > 0 ? mappingName : name);
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
        if ((fieldKind == HtmlRenderFormFieldKind.Text || fieldKind == HtmlRenderFormFieldKind.Choice)
            && !CanPreserveInteractiveFieldTypography(style.Font)) {
            ReportFormFieldTypographyFallback(source, style.Font);
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

        IElement[] allControls = _document.QuerySelectorAll("input,textarea,select")
            .Where(IsSupportedInteractiveFormControl)
            .ToArray();
        IElement[] controls = allControls
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
        if (_options.Mode != HtmlRenderMode.Paged) return;
        foreach (IElement element in allControls.Where(IsInRepeatedPageContent)) {
            string authoredName = element.GetAttribute("name") ?? string.Empty;
            string key = HtmlRenderStyleResolver.DescribeSource(element) + "\n" + (authoredName.Length > 0 ? authoredName : "repeated-page-control");
            _staticRepeatedControlGroupKeys[element] = key;
            if (NormalizeInputType(element) != "radio" || authoredName.Length == 0) continue;
            IElement? owner = HtmlFormControlSemantics.ResolveFormOwner(element);
            foreach (IElement member in radios.Where(candidate =>
                         ReferenceEquals(HtmlFormControlSemantics.ResolveFormOwner(candidate), owner)
                         && string.Equals(candidate.GetAttribute("name"), authoredName, StringComparison.Ordinal))) {
                _staticRepeatedControlGroupKeys[member] = key;
            }
        }
    }

    private bool IsInRepeatedPageContent(IElement element) {
        for (IElement? current = element; current != null; current = current.ParentElement) {
            if (current.LocalName == "thead" || current.LocalName == "tfoot") return true;
            if (_styleResolver.Resolve(current, _options.PageWidth).Position == "fixed") return true;
        }
        return false;
    }

    private static bool IsSupportedInteractiveFormControl(IElement element) {
        if (element.LocalName == "textarea" || element.LocalName == "select") return true;
        if (element.LocalName != "input") return false;
        string type = NormalizeInputType(element);
        return type == "checkbox" || type == "radio" || IsInteractiveTextInputType(type);
    }

    private string ResolveRadioFieldName(IElement element, string mappingName, string partialName, int nodeId) {
        string owner = ResolveFormOwnerKey(element);
        string key = owner + "\n" + mappingName;
        if (_radioFieldNames.TryGetValue(key, out string? name)) return name;
        name = ResolveUniqueFormFieldName(element, partialName, nodeId);
        _radioFieldNames[key] = name;
        return name;
    }

    private static string NormalizePdfPartialFieldName(string name) => name.Replace('.', '-');

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
        string value = NormalizeControlText(HtmlFormControlSemantics.GetValues(element).FirstOrDefault());
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

    private static bool IsInputSizeApplicable(string type) =>
        type == "text"
        || type == "search"
        || type == "tel"
        || type == "url"
        || type == "email"
        || type == "password";

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
        value.Replace("\r\n", "\n").Replace('\r', '\n');

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
