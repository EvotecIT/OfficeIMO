using System;
using System.Collections;
using System.Globalization;
using System.Linq;
using System.Management.Automation;
using OfficeIMO.Excel;

namespace PSWriteOffice.Services.Excel;

internal static class ExcelColumnFormatPlanService
{
    public static ExcelColumnFormatPlan? Build(
        Hashtable? columnFormat,
        string[]? textColumn,
        string[]? numberColumn,
        string[]? integerColumn,
        string[]? percentColumn,
        string[]? currencyColumn,
        string[]? dateColumn,
        string[]? dateTimeColumn,
        int formatDecimals,
        string? formatCultureName)
    {
        if (formatDecimals < 0)
        {
            throw new ArgumentOutOfRangeException(nameof(formatDecimals), "FormatDecimals must be zero or greater.");
        }

        var plan = new ExcelColumnFormatPlan();
        var culture = ResolveCulture(formatCultureName);
        AddPresetColumns(plan, textColumn, ExcelNumberPreset.Text, decimals: 0, culture: null);
        AddPresetColumns(plan, numberColumn, ExcelNumberPreset.Decimal, formatDecimals, culture: null);
        AddPresetColumns(plan, integerColumn, ExcelNumberPreset.Integer, decimals: 0, culture: null);
        AddPresetColumns(plan, percentColumn, ExcelNumberPreset.Percent, formatDecimals, culture: null);
        AddPresetColumns(plan, currencyColumn, ExcelNumberPreset.Currency, formatDecimals, culture);
        AddPresetColumns(plan, dateColumn, ExcelNumberPreset.DateShort, decimals: 0, culture: null);
        AddPresetColumns(plan, dateTimeColumn, ExcelNumberPreset.DateTime, decimals: 0, culture: null);

        if (columnFormat != null)
        {
            foreach (DictionaryEntry entry in columnFormat)
            {
                var header = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(header))
                {
                    continue;
                }

                AddColumnFormatEntry(plan, header!, entry.Value, formatDecimals, culture);
            }
        }

        return plan.Count == 0 ? null : plan;
    }

    private static void AddPresetColumns(ExcelColumnFormatPlan plan, string[]? headers, ExcelNumberPreset preset, int decimals, CultureInfo? culture)
    {
        if (headers == null)
        {
            return;
        }

        foreach (var header in headers.Where(static value => !string.IsNullOrWhiteSpace(value)))
        {
            plan.Add(header.Trim(), preset, decimals, culture);
        }
    }

    private static void AddColumnFormatEntry(ExcelColumnFormatPlan plan, string header, object? value, int formatDecimals, CultureInfo? defaultCulture)
    {
        if (value is PSObject psObject)
        {
            value = psObject.BaseObject;
        }

        if (value is IDictionary dictionary)
        {
            AddDictionaryColumnFormat(plan, header, dictionary, formatDecimals, defaultCulture);
            return;
        }

        var text = Convert.ToString(value, CultureInfo.InvariantCulture);
        if (string.IsNullOrWhiteSpace(text))
        {
            throw new PSArgumentException($"Column format for '{header}' cannot be empty.");
        }

        if (TryResolveNumberPreset(text!, out var resolvedPreset))
        {
            plan.Add(header, resolvedPreset, formatDecimals, defaultCulture);
        }
        else
        {
            plan.AddFormat(header, text!);
        }
    }

    private static void AddDictionaryColumnFormat(
        ExcelColumnFormatPlan plan,
        string header,
        IDictionary dictionary,
        int formatDecimals,
        CultureInfo? defaultCulture)
    {
        var style = GetDictionaryString(dictionary, "Style", "Preset");
        var numberFormat = GetDictionaryString(dictionary, "NumberFormat", "Format", "FormatCode");
        var decimals = GetDictionaryInt(dictionary, "Decimals") ?? formatDecimals;
        var cultureName = GetDictionaryString(dictionary, "CultureName", "Culture");
        var culture = string.IsNullOrWhiteSpace(cultureName) ? defaultCulture : CultureInfo.GetCultureInfo(cultureName!);
        var includeHeader = GetDictionaryBool(dictionary, "IncludeHeader") ?? false;
        var autoFit = GetDictionaryBool(dictionary, "AutoFit") ?? false;

        if (!string.IsNullOrWhiteSpace(style))
        {
            if (IsCustomFormatStyle(style!))
            {
                if (string.IsNullOrWhiteSpace(numberFormat))
                {
                    throw new PSArgumentException($"Column format for '{header}' uses style '{style}' but does not provide NumberFormat or Format.");
                }

                plan.AddFormat(header, numberFormat!, includeHeader, autoFit);
                return;
            }

            if (TryResolveNumberPreset(style!, out var preset))
            {
                plan.Add(header, preset, decimals, culture, includeHeader, autoFit);
                return;
            }
        }

        if (!string.IsNullOrWhiteSpace(numberFormat))
        {
            plan.AddFormat(header, numberFormat!, includeHeader, autoFit);
            return;
        }

        throw new PSArgumentException($"Column format for '{header}' must provide Style, Preset, NumberFormat, or Format.");
    }

    private static CultureInfo? ResolveCulture(string? cultureName)
    {
        return string.IsNullOrWhiteSpace(cultureName) ? null : CultureInfo.GetCultureInfo(cultureName!);
    }

    private static bool TryResolveNumberPreset(string text, out ExcelNumberPreset preset)
    {
        var normalized = text.Trim().Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty);
        if (int.TryParse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
        {
            preset = default;
            return false;
        }

        switch (normalized.ToUpperInvariant())
        {
            case "NUMBER":
            case "DECIMAL":
                preset = ExcelNumberPreset.Decimal;
                return true;
            case "DATE":
            case "DATESHORT":
                preset = ExcelNumberPreset.DateShort;
                return true;
            case "DATELONG":
                preset = ExcelNumberPreset.DateLong;
                return true;
            case "DATETIME":
                preset = ExcelNumberPreset.DateTime;
                return true;
            case "TIME":
                preset = ExcelNumberPreset.Time;
                return true;
            case "DURATION":
            case "DURATIONHOURS":
            case "ELAPSED":
                preset = ExcelNumberPreset.DurationHours;
                return true;
            case "TEXT":
                preset = ExcelNumberPreset.Text;
                return true;
        }

        return Enum.TryParse(normalized, ignoreCase: true, out preset);
    }

    private static bool IsCustomFormatStyle(string style)
    {
        var normalized = style.Trim().Replace("-", string.Empty).Replace("_", string.Empty).Replace(" ", string.Empty);
        return string.Equals(normalized, "NumberFormat", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalized, "Custom", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(normalized, "Format", StringComparison.OrdinalIgnoreCase);
    }

    private static string? GetDictionaryString(IDictionary dictionary, params string[] keys)
    {
        var value = GetDictionaryValue(dictionary, keys);
        return value == null ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
    }

    private static int? GetDictionaryInt(IDictionary dictionary, params string[] keys)
    {
        var value = GetDictionaryValue(dictionary, keys);
        if (value == null)
        {
            return null;
        }

        if (value is int integer)
        {
            return integer;
        }

        var text = Convert.ToString(value, CultureInfo.InvariantCulture);
        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var parsed) ? parsed : null;
    }

    private static bool? GetDictionaryBool(IDictionary dictionary, params string[] keys)
    {
        var value = GetDictionaryValue(dictionary, keys);
        if (value == null)
        {
            return null;
        }

        if (value is bool boolean)
        {
            return boolean;
        }

        if (value is SwitchParameter switchParameter)
        {
            return switchParameter.IsPresent;
        }

        var text = Convert.ToString(value, CultureInfo.InvariantCulture);
        return bool.TryParse(text, out var parsed) ? parsed : null;
    }

    private static object? GetDictionaryValue(IDictionary dictionary, params string[] keys)
    {
        foreach (DictionaryEntry entry in dictionary)
        {
            var key = Convert.ToString(entry.Key, CultureInfo.InvariantCulture);
            if (key == null)
            {
                continue;
            }

            foreach (var requested in keys)
            {
                if (string.Equals(key, requested, StringComparison.OrdinalIgnoreCase))
                {
                    return entry.Value;
                }
            }
        }

        return null;
    }
}
