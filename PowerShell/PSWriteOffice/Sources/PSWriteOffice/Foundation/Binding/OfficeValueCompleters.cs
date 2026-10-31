using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Management.Automation;
using System.Management.Automation.Language;
using System.Reflection;
using OfficeIMO;
using OfficeIMO.Drawing;

namespace PSWriteOffice.Foundation;

/// <summary>Completes representative invariant lengths.</summary>
public sealed class OfficeLengthCompleter : IArgumentCompleter {
    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(string commandName, string parameterName, string wordToComplete, CommandAst commandAst, IDictionary fakeBoundParameters) =>
        OfficeCompletion.Match(new[] { "12pt", "5cm", "1in", "10mm", "96px", "914400emu" }, wordToComplete);
}
/// <summary>Completes CSS color names and semantic theme roles.</summary>
public sealed class OfficeColorCompleter : IArgumentCompleter {
    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(string commandName, string parameterName, string wordToComplete, CommandAst commandAst, IDictionary fakeBoundParameters) =>
        OfficeCompletion.Match(typeof(OfficeColor).GetProperties(BindingFlags.Public | BindingFlags.Static)
            .Where(p => p.PropertyType == typeof(OfficeColor)).Select(p => p.Name), wordToComplete);
}
/// <summary>Completes the built-in shared themes and template paths.</summary>
public sealed class OfficeThemeCompleter : IArgumentCompleter {
    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(string commandName, string parameterName, string wordToComplete, CommandAst commandAst, IDictionary fakeBoundParameters) =>
        OfficeCompletion.Match(new[] { "Modern", "Classic", "Minimal", "Boardroom", "Dark" }, wordToComplete)
            .Concat(CompletionCompleters.CompleteFilename(wordToComplete));
}
/// <summary>Completes the shared format names.</summary>
public sealed class OfficeFormatCompleter : IArgumentCompleter {
    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(string commandName, string parameterName, string wordToComplete, CommandAst commandAst, IDictionary fakeBoundParameters) =>
        OfficeCompletion.Match(Enum.GetNames(typeof(OfficeDocumentFormat)).Where(n => n != "Unknown"), wordToComplete);
}
/// <summary>Completes worksheet names from an explicitly bound document.</summary>
public sealed class OfficeSheetCompleter : IArgumentCompleter {
    /// <inheritdoc />
    public IEnumerable<CompletionResult> CompleteArgument(string commandName, string parameterName, string wordToComplete, CommandAst commandAst, IDictionary fakeBoundParameters) {
        var document = fakeBoundParameters.Contains("Document") ? OfficeArgumentValue.Unwrap(fakeBoundParameters["Document"]) : OfficeDocumentContext.Find<OfficeIMO.Excel.ExcelDocument>();
        return document is OfficeIMO.Excel.ExcelDocument excel
            ? OfficeCompletion.Match(excel.Sheets.Select(s => s.Name), wordToComplete, quote: true)
            : Array.Empty<CompletionResult>();
    }
}
internal static class OfficeCompletion {
    internal static IEnumerable<CompletionResult> Match(IEnumerable<string> candidates, string word, bool quote = false) => candidates
        .Where(n => n.StartsWith(word.Trim('\'', '"'), StringComparison.OrdinalIgnoreCase))
        .Select(n => new CompletionResult(quote ? "'" + n.Replace("'", "''") + "'" : n, n, CompletionResultType.ParameterValue, n));
}
