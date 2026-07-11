namespace OfficeIMO.Markup;

public sealed partial class OfficeMarkupCSharpEmitter {
    private static string Describe(OfficeMarkupBlock block) {
        switch (block) {
            case OfficeMarkupHeadingBlock heading:
                return heading.Text;
            case OfficeMarkupParagraphBlock paragraph:
                return paragraph.Text;
            case OfficeMarkupImageBlock image:
                return image.Source;
            case OfficeMarkupCodeBlock code:
                return code.Language;
            case OfficeMarkupChartBlock chart:
                return chart.Title ?? chart.Source ?? chart.ChartType;
            case OfficeMarkupTextBoxBlock textBox:
                return textBox.Text;
            case OfficeMarkupColumnBlock column:
                return column.ColumnKind;
            case OfficeMarkupCardBlock card:
                return card.Title ?? card.Body;
            case OfficeMarkupExtensionBlock extension:
                return extension.Command;
            default:
                return block.Kind.ToString();
        }
    }

    private static string ToPascalIdentifier(string value) {
        var parts = (value ?? string.Empty).Split(new[] { '-', '_', ' ', '=' }, StringSplitOptions.RemoveEmptyEntries);
        var sb = new StringBuilder();
        foreach (var part in parts) {
            if (part.Length == 0) {
                continue;
            }

            sb.Append(char.ToUpperInvariant(part[0]));
            if (part.Length > 1) {
                sb.Append(part.Substring(1));
            }
        }

        return sb.Length == 0 ? "None" : sb.ToString();
    }

    private static string ToPowerPointChartKind(string chartType) =>
        OfficeMarkupChartKindResolver.Resolve(chartType).ToString();

    private static string ToExcelChartType(string chartType) {
        var normalized = NormalizeToken(chartType);
        return normalized switch {
            "line" => "Line",
            "bar" => "BarClustered",
            "stackedbar" => "BarStacked",
            "stackedcolumn" => "ColumnStacked",
            "pie" => "Pie",
            "doughnut" or "donut" => "Doughnut",
            "scatter" => "Scatter",
            "area" => "Area",
            _ => "ColumnClustered"
        };
    }

    private static string NormalizeToken(string value) =>
        new string((value ?? string.Empty).Where(char.IsLetterOrDigit).Select(char.ToLowerInvariant).ToArray());

    private static void EmitPlacementComment(OfficeMarkupPlacement? placement, StringBuilder sb) {
        if (placement == null || !placement.HasValue) {
            return;
        }

        sb.AppendLine($"// Placement: x={CsString(placement.X ?? string.Empty)}, y={CsString(placement.Y ?? string.Empty)}, w={CsString(placement.Width ?? string.Empty)}, h={CsString(placement.Height ?? string.Empty)}");
    }

    private static void EmitComment(StringBuilder sb, string text) {
        foreach (var line in (text ?? string.Empty).Replace("\r\n", "\n").Split('\n')) {
            sb.AppendLine($"// {line}");
        }
    }

    private static (string SheetExpression, string LocalReference) ResolveWorkbookTarget(string? explicitSheet, string? reference) {
        if (TrySplitSheetQualifiedReference(reference, out var sheetName, out var localReference)) {
            return ($"GetOrAddSheet({CsString(sheetName)})", localReference);
        }

        if (!string.IsNullOrWhiteSpace(explicitSheet)) {
            return ($"GetOrAddSheet({CsString(explicitSheet!.Trim())})", (reference ?? string.Empty).Trim());
        }

        return ($"GetOrAddSheet({CsString("Sheet1")})", (reference ?? string.Empty).Trim());
    }

    private static bool TrySplitSheetQualifiedReference(string? reference, out string sheetName, out string localReference) {
        sheetName = string.Empty;
        localReference = string.Empty;
        if (string.IsNullOrWhiteSpace(reference)) {
            return false;
        }

        var value = reference!.Trim();
        var bangIndex = value.LastIndexOf('!');
        if (bangIndex <= 0 || bangIndex >= value.Length - 1) {
            return false;
        }

        sheetName = value.Substring(0, bangIndex).Trim().Trim('\'').Replace("''", "'");
        localReference = value.Substring(bangIndex + 1).Trim();
        return !string.IsNullOrWhiteSpace(sheetName) && !string.IsNullOrWhiteSpace(localReference);
    }

    private static bool TryParseCellAddress(string? address, out int row, out int column) {
        row = 1;
        column = 1;
        if (string.IsNullOrWhiteSpace(address)) {
            return false;
        }

        var match = System.Text.RegularExpressions.Regex.Match(address!.Trim(), @"^\$?([A-Za-z]+)\$?(\d+)");
        if (!match.Success) {
            return false;
        }

        column = 0;
        foreach (var character in match.Groups[1]!.Value.ToUpperInvariant()) {
            column = (column * 26) + character - 'A' + 1;
        }

        row = int.Parse(match.Groups[2]!.Value, System.Globalization.CultureInfo.InvariantCulture);
        return row > 0 && column > 0;
    }

    private static bool TryParseInt(string? value, out int result) =>
        int.TryParse(value, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out result);

    private static string NumericLiteral(string? value) {
        if (double.TryParse(value, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out var number)) {
            return number.ToString("0.###############", System.Globalization.CultureInfo.InvariantCulture) + "d";
        }

        return "0d";
    }

    private static string FormatList(OfficeMarkupListBlock list) {
        var lines = list.Items.Select((item, index) => list.Ordered
            ? $"{list.Start + index}. {item.Text}"
            : $"- {item.Text}");
        return string.Join(Environment.NewLine, lines);
    }

    private static string CsString(string value) {
        return "@\"" + (value ?? string.Empty).Replace("\"", "\"\"") + "\"";
    }

    private static string? GetAttribute(IDictionary<string, string> attributes, params string[] names) {
        foreach (var name in names) {
            if (attributes.TryGetValue(name, out var value) && !string.IsNullOrWhiteSpace(value)) {
                return value.Trim();
            }
        }

        return null;
    }

    private static bool IsTruthy(string value) {
        var normalized = NormalizeToken(value);
        return normalized is not ("false" or "no" or "off" or "none" or "hidden" or "0");
    }

    private static bool TryGetLegendPositionIdentifier(string value, out string identifier) {
        identifier = NormalizeToken(value) switch {
            "left" => "Left",
            "right" => "Right",
            "top" => "Top",
            "bottom" => "Bottom",
            "corner" or "topright" => "TopRight",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static bool TryGetDataLabelPositionIdentifier(string? value, out string identifier) {
        identifier = NormalizeToken(value ?? string.Empty) switch {
            "center" => "Center",
            "insideend" => "InsideEnd",
            "insidebase" => "InsideBase",
            "outsideend" => "OutsideEnd",
            "bestfit" => "BestFit",
            "left" => "Left",
            "right" => "Right",
            "top" => "Top",
            "bottom" => "Bottom",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static bool TryGetHorizontalAlignmentIdentifier(string? value, out string identifier) {
        identifier = NormalizeToken(value ?? string.Empty) switch {
            "general" => "General",
            "left" => "Left",
            "center" or "centre" => "Center",
            "centercontinuous" or "centeracross" or "centeracrossselection" => "CenterContinuous",
            "right" => "Right",
            "fill" => "Fill",
            "justify" => "Justify",
            "distributed" => "Distributed",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static bool TryGetVerticalAlignmentIdentifier(string? value, out string identifier) {
        identifier = NormalizeToken(value ?? string.Empty) switch {
            "top" => "Top",
            "middle" or "center" or "centre" => "Center",
            "bottom" => "Bottom",
            "justify" => "Justify",
            "distributed" => "Distributed",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static bool TryGetBorderStyleIdentifier(string? value, out string identifier) {
        identifier = NormalizeToken(value ?? string.Empty) switch {
            "true" or "yes" or "on" or "1" or "thin" => "Thin",
            "medium" => "Medium",
            "thick" => "Thick",
            "dashed" => "Dashed",
            "dotted" => "Dotted",
            "double" => "Double",
            "hair" => "Hair",
            "dashdot" => "DashDot",
            "dashdotdot" => "DashDotDot",
            "mediumdashed" => "MediumDashed",
            "mediumdashdot" => "MediumDashDot",
            "mediumdashdotdot" => "MediumDashDotDot",
            "slantdashdot" => "SlantDashDot",
            _ => string.Empty
        };
        return identifier.Length > 0;
    }

    private static IEnumerable<(int Row, int Column)> EnumerateTargetCells(string target) {
        if (string.IsNullOrWhiteSpace(target)) {
            yield break;
        }

        var parts = target.Split(new[] { ':' }, 2, StringSplitOptions.None)
            .Select(part => part.Trim())
            .ToArray();
        if (parts.Length == 1) {
            if (TryParseCellAddress(parts[0], out var singleRow, out var singleColumn)) {
                yield return (singleRow, singleColumn);
            }

            yield break;
        }

        if (!TryParseCellAddress(parts[0], out var startRow, out var startColumn) ||
            !TryParseCellAddress(parts[1], out var endRow, out var endColumn)) {
            yield break;
        }

        if (endRow < startRow) {
            (startRow, endRow) = (endRow, startRow);
        }

        if (endColumn < startColumn) {
            (startColumn, endColumn) = (endColumn, startColumn);
        }

        for (var row = startRow; row <= endRow; row++) {
            for (var column = startColumn; column <= endColumn; column++) {
                yield return (row, column);
            }
        }
    }

    private static string BoolLiteral(bool value) => value ? "true" : "false";
}
