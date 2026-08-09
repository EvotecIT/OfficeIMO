using OfficeIMO.Excel;
using OfficeIMO.OpenDocument;

namespace OfficeIMO.Excel.OpenDocument;

public static partial class ExcelOpenDocumentConversionExtensions {
    private sealed class NamedRangeConversionEntry {
        internal NamedRangeConversionEntry(string outputName, string address) {
            OutputName = outputName;
            Address = address;
        }

        internal string OutputName { get; }
        internal string Address { get; }
    }

    private sealed class NamedRangeConversionPlan {
        private readonly IReadOnlyDictionary<string, string> _globalNames;
        private readonly IReadOnlyDictionary<string, Dictionary<string, string>> _localNames;

        internal NamedRangeConversionPlan(
            IReadOnlyList<NamedRangeConversionEntry> entries,
            IReadOnlyDictionary<string, string> globalNames,
            IReadOnlyDictionary<string, Dictionary<string, string>> localNames,
            int builtInCount,
            int unsupportedExpressionCount,
            int disambiguatedCount) {
            Entries = entries;
            _globalNames = globalNames;
            _localNames = localNames;
            BuiltInCount = builtInCount;
            UnsupportedExpressionCount = unsupportedExpressionCount;
            DisambiguatedCount = disambiguatedCount;
        }

        internal IReadOnlyList<NamedRangeConversionEntry> Entries { get; }
        internal int BuiltInCount { get; }
        internal int UnsupportedExpressionCount { get; }
        internal int DisambiguatedCount { get; }

        internal string RewriteFormula(string formula, string worksheetName) {
            ExcelFormulaSyntaxTree syntax = ExcelFormulaSyntaxTree.Parse(formula);
            return syntax.RewriteNames(authoredName => ResolveName(authoredName, worksheetName));
        }

        private string ResolveName(string authoredName, string worksheetName) {
            if (TrySplitQualifiedName(authoredName, out string? qualifiedSheet, out string? localName)) {
                return _localNames.TryGetValue(qualifiedSheet!, out Dictionary<string, string>? qualifiedNames)
                    && qualifiedNames.TryGetValue(localName!, out string? qualifiedOutput)
                    ? qualifiedOutput
                    : authoredName;
            }
            if (_localNames.TryGetValue(worksheetName, out Dictionary<string, string>? localNames)
                && localNames.TryGetValue(authoredName, out string? localOutput)) return localOutput;
            return _globalNames.TryGetValue(authoredName, out string? globalOutput) ? globalOutput : authoredName;
        }
    }

    private static NamedRangeConversionPlan BuildNamedRangeConversionPlan(
        IReadOnlyList<ExcelNamedRangeSnapshot> namedRanges) {
        var entries = new List<NamedRangeConversionEntry>();
        var globalNames = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var localNames = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
        var usedOutputNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int builtInCount = 0;
        int unsupportedExpressionCount = 0;
        int disambiguatedCount = 0;

        foreach (ExcelNamedRangeSnapshot named in namedRanges) {
            if (named.IsBuiltIn) {
                builtInCount++;
                continue;
            }
            string address = SpreadsheetAddressConverter.ExcelRangeToOpenAddress(named.ReferenceA1, named.SheetName);
            if (address.Length == 0) {
                unsupportedExpressionCount++;
                continue;
            }

            string outputName = named.Name;
            if (!usedOutputNames.Add(outputName)) {
                outputName = CreateUniqueNamedRangeName(named.Name, named.SheetName, usedOutputNames);
                disambiguatedCount++;
            }
            entries.Add(new NamedRangeConversionEntry(outputName, address));

            Dictionary<string, string> scope;
            if (named.SheetName == null) {
                scope = globalNames;
            } else {
                if (!localNames.TryGetValue(named.SheetName, out Dictionary<string, string>? existingScope)) {
                    existingScope = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    localNames.Add(named.SheetName, existingScope);
                }
                scope = existingScope;
            }
            scope[named.Name] = outputName;
        }

        return new NamedRangeConversionPlan(
            entries,
            globalNames,
            localNames,
            builtInCount,
            unsupportedExpressionCount,
            disambiguatedCount);
    }

    private static bool TrySplitQualifiedName(string authoredName, out string? sheetName, out string? localName) {
        sheetName = null;
        localName = null;
        int separator = authoredName.LastIndexOf('!');
        if (separator <= 0 || separator == authoredName.Length - 1) return false;
        string qualifier = authoredName.Substring(0, separator);
        if (qualifier.Length >= 2 && qualifier[0] == '\'' && qualifier[qualifier.Length - 1] == '\'') {
            qualifier = qualifier.Substring(1, qualifier.Length - 2).Replace("''", "'");
        } else if (qualifier.IndexOf('[') >= 0 || qualifier.IndexOf(']') >= 0) {
            return false;
        }
        sheetName = qualifier;
        localName = authoredName.Substring(separator + 1);
        return true;
    }

    private static string CreateUniqueNamedRangeName(string name, string? sheetName, HashSet<string> usedNames) {
        string suffix = new string((sheetName ?? "Sheet").Select(character => char.IsLetterOrDigit(character) ? character : '_').ToArray());
        if (suffix.Length == 0) suffix = "Sheet";
        string candidate = name + "__" + suffix;
        int index = 2;
        while (!usedNames.Add(candidate)) candidate = name + "__" + suffix + "_" + index++.ToString(CultureInfo.InvariantCulture);
        return candidate;
    }

    private sealed class OdsNamedRangeConversionPlan {
        private readonly IReadOnlyDictionary<string, string> _outputNames;

        internal OdsNamedRangeConversionPlan(
            IReadOnlyList<NamedRangeConversionEntry> entries,
            IReadOnlyDictionary<string, string> outputNames,
            int renamedCount) {
            Entries = entries;
            _outputNames = outputNames;
            RenamedCount = renamedCount;
        }

        internal IReadOnlyList<NamedRangeConversionEntry> Entries { get; }
        internal int RenamedCount { get; }

        internal string RewriteFormula(string formula) => ExcelFormulaSyntaxTree.Parse(formula)
            .RewriteNames(name => _outputNames.TryGetValue(name, out string? outputName) ? outputName : name);

        internal bool TryResolveName(string sourceName, out string outputName) =>
            _outputNames.TryGetValue(sourceName, out outputName!);
    }

    private static OdsNamedRangeConversionPlan BuildOdsNamedRangeConversionPlan(
        IReadOnlyList<OdsNamedRange> namedRanges) {
        var entries = new List<NamedRangeConversionEntry>();
        var outputNames = new Dictionary<string, string>(StringComparer.Ordinal);
        var usedOutputNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int renamedCount = 0;

        foreach (OdsNamedRange named in namedRanges) {
            string address = SpreadsheetAddressConverter.OpenAddressToExcel(named.CellRangeAddress);
            if (address.Length == 0) continue;

            string outputName = ExcelDocument.NormalizeDefinedName(named.Name);
            if (!usedOutputNames.Add(outputName)) outputName = CreateUniqueExcelDefinedName(outputName, usedOutputNames);
            if (!string.Equals(outputName, named.Name, StringComparison.Ordinal)) renamedCount++;
            entries.Add(new NamedRangeConversionEntry(outputName, address));
            outputNames[named.Name] = outputName;
        }

        return new OdsNamedRangeConversionPlan(entries, outputNames, renamedCount);
    }

    private static string CreateUniqueExcelDefinedName(string normalizedName, HashSet<string> usedNames) {
        for (int index = 2; ; index++) {
            string suffix = "_" + index.ToString(CultureInfo.InvariantCulture);
            int prefixLength = Math.Min(normalizedName.Length, ExcelDocument.MaximumDefinedNameLength - suffix.Length);
            string candidate = normalizedName.Substring(0, prefixLength) + suffix;
            if (usedNames.Add(candidate)) return candidate;
        }
    }
}