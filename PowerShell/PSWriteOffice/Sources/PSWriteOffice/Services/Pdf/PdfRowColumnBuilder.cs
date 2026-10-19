using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Services.Pdf;

internal static class PdfRowColumnBuilder
{
    internal static void AddContent(PdfContentBuilder column, object specification)
    {
        if (TryGetValue(specification, out var content, "Content", "Blocks"))
        {
            if (IsEnumerableContent(content))
            {
                foreach (var item in ((IEnumerable)content!).Cast<object>())
                {
                    AddItem(column, item);
                }
            }
            else
            {
                AddItem(column, content!);
            }

            return;
        }

        if (!HasExplicitType(specification) &&
            TryGetValue(specification, out content, "Items") &&
            IsEnumerableContent(content))
        {
            foreach (var item in ((IEnumerable)content!).Cast<object>())
            {
                AddItem(column, item);
            }

            return;
        }

        AddShorthand(column, specification);
    }

    private static void AddItem(PdfContentBuilder column, object specification)
    {
        var type = GetString(specification, "Type", "Kind", "Block");
        if (string.IsNullOrWhiteSpace(type))
        {
            AddShorthand(column, specification);
            return;
        }

        switch (Normalize(type!))
        {
            case "h1":
            case "h2":
            case "h3":
            case "heading":
                AddHeading(column, specification);
                break;
            case "paragraph":
            case "text":
                AddParagraph(column, specification);
                break;
            case "panel":
            case "callout":
                column.PanelParagraph(p => p.Text(GetRequiredString(specification, "Text", "Panel")), align: GetAlign(specification), defaultColor: GetColor(specification, "Color"));
                break;
            case "list":
            case "bullets":
            case "numbered":
                AddList(column, specification, string.Equals(Normalize(type!), "numbered", StringComparison.Ordinal));
                break;
            case "table":
                column.Table(GetTableRows(specification), GetAlign(specification), PdfTableStyleBuilder.CreateFromSpecification(specification));
                break;
            case "rule":
            case "hr":
            case "horizontalrule":
                column.HR(GetDouble(specification, "Thickness"), GetColor(specification, "Color"), GetDouble(specification, "SpacingBefore"), GetDouble(specification, "SpacingAfter"));
                break;
            case "spacer":
            case "space":
                column.Spacer(GetDouble(specification, "Height", "Size") ?? 12D);
                break;
            case "bookmark":
                column.Bookmark(GetRequiredString(specification, "Name", "Bookmark"));
                break;
            default:
                throw new PSArgumentException($"Unsupported PDF row content type '{type}'.");
        }
    }

    private static void AddShorthand(PdfContentBuilder column, object specification)
    {
        if (TryGetString(specification, out var bookmark, "Bookmark", "Name"))
        {
            column.Bookmark(bookmark);
        }

        if (TryGetString(specification, out _, "Heading", "Title"))
        {
            AddHeading(column, specification);
        }

        if (TryGetRuns(specification, out var runs))
        {
            column.Paragraph(p => PdfRichTextRunBuilder.ApplyRuns(p, runs), GetAlign(specification), GetColor(specification, "Color"));
        }
        else if (TryGetString(specification, out var paragraph, "Paragraph", "Text"))
        {
            column.Paragraph(p => p.Text(paragraph), GetAlign(specification), GetColor(specification, "Color"));
        }

        if (TryGetString(specification, out var panel, "Panel", "Callout"))
        {
            column.PanelParagraph(p => p.Text(panel), align: GetAlign(specification), defaultColor: GetColor(specification, "PanelColor", "Color"));
        }

        if (TryGetValue(specification, out var listValue, "List", "Bullets", "Items"))
        {
            AddList(column, specification, GetBool(specification, "Numbered"), listValue);
        }

        if (TryGetValue(specification, out _, "Table", "Rows", "InputObject"))
        {
            column.Table(GetTableRows(specification), GetAlign(specification), PdfTableStyleBuilder.CreateFromSpecification(specification));
        }

        if (GetBool(specification, "Rule", "HorizontalRule", "Hr"))
        {
            column.HR(GetDouble(specification, "Thickness"), GetColor(specification, "RuleColor", "Color"), GetDouble(specification, "SpacingBefore"), GetDouble(specification, "SpacingAfter"));
        }

        if (TryGetValue(specification, out _, "Spacer", "Space", "Height"))
        {
            column.Spacer(GetDouble(specification, "Spacer", "Space", "Height") ?? 12D);
        }
    }

    private static void AddHeading(PdfContentBuilder column, object specification)
    {
        var text = GetString(specification, "Text", "Heading", "Title") ?? string.Empty;
        var level = GetInt(specification, "Level", "HeadingLevel") ?? GetLevelFromType(specification);
        var align = GetAlign(specification, "HeadingAlign", "Align");
        var color = GetColor(specification, "HeadingColor", "Color");

        switch (Math.Max(1, Math.Min(3, level)))
        {
            case 1:
                column.H1(text, align, color);
                break;
            case 2:
                column.H2(text, align, color);
                break;
            default:
                column.H3(text, align, color);
                break;
        }
    }

    private static void AddParagraph(PdfContentBuilder column, object specification)
    {
        if (TryGetRuns(specification, out var runs))
        {
            column.Paragraph(p => PdfRichTextRunBuilder.ApplyRuns(p, runs), GetAlign(specification), GetColor(specification, "Color"));
            return;
        }

        column.Paragraph(p => p.Text(GetRequiredString(specification, "Text", "Paragraph")), GetAlign(specification), GetColor(specification, "Color"));
    }

    private static void AddList(PdfContentBuilder column, object specification, bool numbered, object? explicitValue = null)
    {
        var items = GetStringArray(explicitValue ?? GetValue(specification, "List", "Bullets", "Items"));
        if (numbered || GetBool(specification, "Numbered"))
        {
            column.Numbered(items, GetAlign(specification), GetColor(specification, "Color"), GetInt(specification, "StartNumber") ?? 1);
        }
        else
        {
            column.Bullets(items, GetAlign(specification), GetColor(specification, "Color"));
        }
    }

    private static string[][] GetTableRows(object specification)
    {
        var value = GetValue(specification, "Table", "Rows", "InputObject")
            ?? throw new PSArgumentException("Table row content requires Table, Rows, or InputObject.");
        var property = GetOptionalStringArray(specification, "Property", "Properties");
        var header = GetOptionalStringArray(specification, "Header", "Headers");

        if (value is IDictionary || value is not IEnumerable || value is string)
        {
            return PdfCommandUtilities.ConvertToTableRows(new[] { value }, property, header);
        }

        if (value is IEnumerable enumerable)
        {
            var items = enumerable.Cast<object>().ToArray();
            if (items.Length == 0)
            {
                throw new PSArgumentException("Provide at least one table row.");
            }

            var rowLike = items[0] is IEnumerable && items[0] is not string && items[0] is not IDictionary;
            return rowLike
                ? PdfCommandUtilities.ConvertDataRows(items, header)
                : PdfCommandUtilities.ConvertToTableRows(items, property, header);
        }

        throw new PSArgumentException("Table content must be an enumerable of objects or row arrays.");
    }

    private static int GetLevelFromType(object specification)
    {
        var type = Normalize(GetString(specification, "Type", "Kind", "Block") ?? string.Empty);
        return type switch
        {
            "h1" => 1,
            "h2" => 2,
            "h3" => 3,
            _ => 1
        };
    }

    internal static double GetWidth(object specification, double defaultWidth)
    {
        return GetDouble(specification, "Width", "WidthPercent", "Percent") ?? defaultWidth;
    }

    private static PdfAlign GetAlign(object specification, params string[] names)
    {
        var value = names.Length == 0 ? GetValue(specification, "Align") : GetValue(specification, names);
        if (value is null)
        {
            return PdfAlign.Left;
        }

        return value is PdfAlign align
            ? align
            : (PdfAlign)Enum.Parse(typeof(PdfAlign), Convert.ToString(value, CultureInfo.InvariantCulture) ?? "Left", ignoreCase: true);
    }

    private static PdfColor? GetColor(object specification, params string[] names)
    {
        return PdfCommandUtilities.ParseColor(GetString(specification, names));
    }

    private static string GetRequiredString(object specification, params string[] names)
    {
        var value = GetString(specification, names);
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new PSArgumentException($"Missing required value: {string.Join(", ", names)}.");
        }

        return value!;
    }

    private static bool TryGetString(object specification, out string value, params string[] names)
    {
        value = GetString(specification, names) ?? string.Empty;
        return !string.IsNullOrWhiteSpace(value);
    }

    private static bool TryGetRuns(object specification, out object[] runs)
    {
        if (TryGetValue(specification, out var value, "Run", "Runs"))
        {
            runs = PdfRichTextRunBuilder.ToRunArray(value);
            return runs.Length > 0;
        }

        runs = Array.Empty<object>();
        return false;
    }

    private static string? GetString(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    private static double? GetDouble(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        return value == null ? null : Convert.ToDouble(value, CultureInfo.InvariantCulture);
    }

    private static int? GetInt(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        return value == null ? null : Convert.ToInt32(value, CultureInfo.InvariantCulture);
    }

    private static bool GetBool(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        return value != null && Convert.ToBoolean(value, CultureInfo.InvariantCulture);
    }

    private static string[]? GetOptionalStringArray(object specification, params string[] names)
    {
        var value = GetValue(specification, names);
        return value == null ? null : GetStringArray(value);
    }

    private static string[] GetStringArray(object? value)
    {
        if (value == null)
        {
            return Array.Empty<string>();
        }

        if (value is string text)
        {
            return new[] { text };
        }

        if (value is IEnumerable enumerable)
        {
            return enumerable.Cast<object?>()
                .Select(item => Convert.ToString(item, CultureInfo.InvariantCulture) ?? string.Empty)
                .ToArray();
        }

        return new[] { Convert.ToString(value, CultureInfo.InvariantCulture) ?? string.Empty };
    }

    private static object? GetValue(object specification, params string[] names)
    {
        TryGetValue(specification, out var value, names);
        return value;
    }

    private static bool TryGetValue(object specification, out object? value, params string[] names)
    {
        if (specification is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                if (names.Any(name => string.Equals(name, key, StringComparison.OrdinalIgnoreCase)))
                {
                    value = entry.Value;
                    return true;
                }
            }
        }

        var psObject = PSObject.AsPSObject(specification);
        foreach (var name in names)
        {
            var property = psObject.Properties
                .Cast<PSPropertyInfo>()
                .FirstOrDefault(candidate => string.Equals(candidate.Name, name, StringComparison.OrdinalIgnoreCase));
            if (property != null)
            {
                value = property.Value;
                return true;
            }
        }

        value = null;
        return false;
    }

    private static bool IsEnumerableContent(object? value)
    {
        return value is IEnumerable && value is not string && value is not IDictionary;
    }

    private static bool HasExplicitType(object specification)
    {
        return !string.IsNullOrWhiteSpace(GetString(specification, "Type", "Kind", "Block"));
    }

    private static string Normalize(string value)
    {
        return value.Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty).ToLowerInvariant();
    }
}
