using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Language;
using System.Reflection;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;

namespace PSWriteOffice.Completion;

internal static class OfficeCompletionResultFactory
{
    internal static IEnumerable<CompletionResult> Complete(IEnumerable<string> values, string wordToComplete, string description)
    {
        var prefix = (wordToComplete ?? string.Empty).TrimStart('\'', '"');
        return values
            .Where(value => value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
            .Select(value => new CompletionResult(value, value, CompletionResultType.ParameterValue, description));
    }
}

/// <summary>Normalizes named Office colors to portable hexadecimal values while preserving caller-supplied hex notation.</summary>
public sealed class OfficeColorArgumentTransformationAttribute : ArgumentTransformationAttribute
{
    /// <summary>Allow the <c>None</c> sentinel in addition to color values.</summary>
    public bool AllowNone { get; set; }

    /// <summary>Normalize shorthand and named colors to Excel-compatible RGB or ARGB notation while preserving legacy eight-digit ARGB values.</summary>
    public bool UseExcelArgb { get; set; }

    /// <inheritdoc />
    public override object Transform(EngineIntrinsics engineIntrinsics, object inputData)
    {
        var value = inputData is PSObject psObject ? psObject.BaseObject : inputData;
        if (value is string text)
        {
            return Normalize(text);
        }

        if (value is IEnumerable sequence)
        {
            return sequence.Cast<object>().Select(item => Normalize(LanguagePrimitives.ConvertTo<string>(item))).ToArray();
        }

        return inputData;
    }

    private string Normalize(string value)
    {
        var trimmed = value.Trim();
        if (AllowNone && string.Equals(trimmed, "None", StringComparison.OrdinalIgnoreCase))
        {
            return "None";
        }

        var hex = trimmed.TrimStart('#');
        if ((hex.Length == 3 || hex.Length == 4 || hex.Length == 6 || hex.Length == 8) && hex.All(Uri.IsHexDigit))
        {
            if (hex.Length == 8 || (!UseExcelArgb && hex.Length == 4))
            {
                return trimmed;
            }

            var parsedHex = OfficeColor.Parse(trimmed);
            return parsedHex.A == byte.MaxValue
                ? "#" + parsedHex.ToRgbHex()
                : "#" + parsedHex.ToArgbHex();
        }

        if (!OfficeColor.TryParse(trimmed, out var color))
        {
            throw new PSArgumentException($"Color must be a known Office color name or a 3, 4, 6, or 8 digit hexadecimal value. Received '{value}'.");
        }

        return color.A == byte.MaxValue ? "#" + color.ToRgbHex() : "#" + color.ToArgbHex();
    }
}

/// <summary>Completes named colors accepted by OfficeIMO while retaining support for hexadecimal values.</summary>
public sealed class OfficeColorArgumentCompleter : IArgumentCompleter
{
    internal static readonly string[] Names = typeof(OfficeColor)
        .GetProperties(BindingFlags.Public | BindingFlags.Static)
        .Where(property => property.PropertyType == typeof(OfficeColor))
        .Select(property => property.Name)
        .Where(name => !string.Equals(name, nameof(OfficeColor.Transparent), StringComparison.OrdinalIgnoreCase))
        .ToArray();

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Named Office color; #RGB, #RRGGBB, and alpha-bearing hex values are also accepted.");
}

/// <summary>Completes background and highlight colors, including the canonical <c>None</c> sentinel.</summary>
public sealed class OfficeHighlightColorArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = OfficeColorArgumentCompleter.Names.Concat(new[] { "None" }).ToArray();

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Named Office color, hexadecimal color, or None to clear the run background or highlight.");
}

/// <summary>Completes known PDF page sizes while retaining the Custom size option.</summary>
public sealed class OfficePdfPageSizeArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = PageSizes.Names.Concat(new[] { "Custom" }).ToArray();

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Known PDF page size; use Custom together with Width and Height.");
}

/// <summary>Completes built-in PSWriteOffice and Word-compatible PDF table styles.</summary>
public sealed class OfficePdfTableStyleArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = new[]
        {
            "Light", "Minimal", "RightAlignedNumbers", "TechnicalDocument", "Compact", "Report"
        }
        .Concat(TableStyles.SupportedWordStyleNames)
        .ToArray();

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Built-in PSWriteOffice or Word-compatible table style.");
}

/// <summary>Completes common Markdown callout kinds without closing the extension-friendly string domain.</summary>
public sealed class OfficeMarkdownCalloutKindArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = { "Note", "Tip", "Important", "Warning", "Caution" };

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Common Markdown callout kind; custom renderer-specific kinds remain accepted.");
}

/// <summary>Completes format-neutral rich-text run kinds.</summary>
public sealed class OfficeTextRunKindArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = { "Text", "LineBreak", "Tab", "Superscript", "Subscript" };

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Portable rich-text run kind; target-specific values remain accepted.");
}

/// <summary>Completes common cross-format underline styles.</summary>
public sealed class OfficeUnderlineStyleArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = { "Single", "Double", "Dotted", "Dash", "Wave", "Words" };

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Common underline style; availability depends on the target format.");
}

/// <summary>Completes common cross-format baseline values.</summary>
public sealed class OfficeTextBaselineArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = { "Normal", "Superscript", "Subscript" };

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Common text baseline; target-specific values remain accepted.");
}

/// <summary>Completes common PDF tab leader values.</summary>
public sealed class OfficeTabLeaderArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = { "None", "Dots", "Hyphens", "Underscores" };

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Common tab leader; target-specific values remain accepted.");
}

/// <summary>Completes common tab alignment values.</summary>
public sealed class OfficeTabAlignmentArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = { "Left", "Center", "Right", "DecimalSeparator" };

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Common tab alignment; target-specific values remain accepted.");
}

/// <summary>Completes canonical line-ending names while allowing APIs that support literal separators to remain open.</summary>
public sealed class OfficeLineEndingArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = { "LF", "CRLF", "CR" };

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Canonical line ending.");
}

/// <summary>Completes built-in Excel pivot table style names while allowing custom workbook styles.</summary>
public sealed class OfficeExcelPivotStyleArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = new[] { "Light", "Medium", "Dark" }
        .SelectMany(family => Enumerable.Range(1, 28).Select(index => $"PivotStyle{family}{index}"))
        .ToArray();

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Built-in Excel pivot table style; custom workbook style names remain accepted.");
}

/// <summary>Completes common document asset kinds while allowing reader-specific kinds.</summary>
public sealed class OfficeDocumentAssetKindArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names = { "Image", "Preview", "Embedded-Object" };

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Common reader asset kind; format-specific kinds remain accepted.");
}

/// <summary>Completes standard HTML referrer-policy tokens.</summary>
public sealed class OfficeReferrerPolicyArgumentCompleter : IArgumentCompleter
{
    private static readonly string[] Names =
    {
        "no-referrer", "no-referrer-when-downgrade", "origin", "origin-when-cross-origin",
        "same-origin", "strict-origin", "strict-origin-when-cross-origin", "unsafe-url"
    };

    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(
        string commandName,
        string parameterName,
        string wordToComplete,
        CommandAst commandAst,
        IDictionary fakeBoundParameters) =>
        OfficeCompletionResultFactory.Complete(Names, wordToComplete, "Standard HTML referrer-policy token.");
}
