using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static readonly Regex TableRnCnReferencePattern = new Regex(
            @"(?<![A-Z0-9_])(?<start>R[1-9][0-9]*C[1-9][0-9]*)(?:\s*:\s*(?<end>R[1-9][0-9]*C[1-9][0-9]*))?(?![A-Z0-9_])",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static readonly Regex ExactTableRnCnAddressPattern = new Regex(
            @"^R(?<row>[1-9][0-9]*)C(?<column>[1-9][0-9]*)$",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);

        private static bool TryResolveRnCnTableCellReferences(
            MutableFieldCandidate candidate,
            string expression,
            out string resolvedExpression,
            out string? diagnostic) {
            resolvedExpression = expression;
            diagnostic = null;

            MatchCollection matches = TableRnCnReferencePattern.Matches(expression);
            if (matches.Count == 0) {
                return true;
            }

            if (!TryGetFieldTable(candidate, out Table? table, out diagnostic) || table == null) {
                return false;
            }

            List<TableRow> rows = table.Elements<TableRow>().ToList();
            var builder = new StringBuilder();
            int startIndex = 0;

            foreach (Match match in matches) {
                if (!TryParseRnCnTableAddress(match.Groups["start"].Value, out TableAddress startAddress)) {
                    diagnostic = $"Formula table RnCn reference {match.Groups["start"].Value} could not be parsed.";
                    return false;
                }

                TableAddress endAddress = startAddress;
                if (match.Groups["end"].Success && !TryParseRnCnTableAddress(match.Groups["end"].Value, out endAddress)) {
                    diagnostic = $"Formula table RnCn reference {match.Groups["end"].Value} could not be parsed.";
                    return false;
                }

                if (!TryResolveTableAddressRange(rows, startAddress, endAddress, out IReadOnlyList<decimal> values, out diagnostic)) {
                    return false;
                }

                if (values.Count == 0) {
                    diagnostic = $"Formula table RnCn reference {match.Value} did not resolve any numeric cells.";
                    return false;
                }

                builder.Append(expression, startIndex, match.Index - startIndex);
                builder.Append(string.Join(", ", values.Select(FormatFormulaValue)));
                startIndex = match.Index + match.Length;
            }

            builder.Append(expression, startIndex, expression.Length - startIndex);
            resolvedExpression = builder.ToString();
            return true;
        }

        private static bool TryParseRnCnTableAddress(string text, out TableAddress address) {
            address = default;
            Match match = ExactTableRnCnAddressPattern.Match(text.Trim());
            if (!match.Success) {
                return false;
            }

            if (!int.TryParse(match.Groups["row"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int oneBasedRow)
                || !int.TryParse(match.Groups["column"].Value, NumberStyles.None, CultureInfo.InvariantCulture, out int oneBasedColumn)) {
                return false;
            }

            try {
                address = new TableAddress(checked(oneBasedColumn - 1), checked(oneBasedRow - 1));
                return true;
            } catch (OverflowException) {
                return false;
            }
        }
    }
}
