using System;
using System.Text;
using System.Text.RegularExpressions;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Owns the bounded lexical grammar used to discover A1 references in formula-bearing workbook metadata.
    /// </summary>
    /// <remarks>
    /// Consumers remain responsible for deciding whether a discovered reference belongs to the edited sheet
    /// and how it should move. Keeping discovery here prevents shared formulas, structural edits, names,
    /// charts, validations, and template operations from drifting into incompatible reference grammars.
    /// </remarks>
    internal static class ExcelFormulaReferenceRewriter {
        private static readonly TimeSpan SharedFormulaRegexTimeout = TimeSpan.FromSeconds(1);
        private static readonly TimeSpan StructuralReferenceRegexTimeout = TimeSpan.FromMilliseconds(200);

        internal static readonly Regex SharedFormulaReferenceRegex = new Regex(
            @"(?<![\p{L}\p{M}\p{N}_\.\\])(?<qualifier>(?:'(?:[^']|'')+'|\[[^\]]+\][\p{L}\p{M}\p{N}_\. ]+|[\p{L}_][\p{L}\p{M}\p{N}_\. ]*:[\p{L}_][\p{L}\p{M}\p{N}_\. ]*|[\p{L}_][\p{L}\p{M}\p{N}_\. ]*)!)?(?:(?<cellStartColumnAbsolute>\$?)(?<cellStartColumn>[A-Za-z]{1,3})(?<cellStartRowAbsolute>\$?)(?<cellStartRow>\d{1,7})(?::(?<cellEndColumnAbsolute>\$?)(?<cellEndColumn>[A-Za-z]{1,3})(?<cellEndRowAbsolute>\$?)(?<cellEndRow>\d{1,7}))?(?<cellSpill>#)?|(?<wholeStartColumnAbsolute>\$?)(?<wholeStartColumn>[A-Za-z]{1,3}):(?<wholeEndColumnAbsolute>\$?)(?<wholeEndColumn>[A-Za-z]{1,3})|(?<wholeStartRowAbsolute>\$?)(?<wholeStartRow>\d{1,7}):(?<wholeEndRowAbsolute>\$?)(?<wholeEndRow>\d{1,7}))(?![\p{L}\p{M}\p{N}_\.]|\()",
            RegexOptions.IgnoreCase | RegexOptions.Compiled | RegexOptions.CultureInvariant,
            SharedFormulaRegexTimeout);

        internal static readonly Regex CellReferenceRegex = new Regex(
            @"(?<![\p{L}\p{N}_\.\]:!\\])(?:(?<sheet>'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_\.]*)!)?(?<colAbs>\$?)(?<col>[A-Za-z]{1,3})(?<rowAbs>\$?)(?<row>\d{1,7})(?=[:),+\-*/^&=<>%# \t\r\n]|$)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant,
            StructuralReferenceRegexTimeout);

        internal static readonly Regex RangeReferenceRegex = new Regex(
            @"(?<![\p{L}\p{N}_\.\]:!\\])(?:(?<sheet>'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_\.]*)!)?(?<startColAbs>\$?)(?<startCol>[A-Za-z]{1,3})(?<startRowAbs>\$?)(?<startRow>\d{1,7}):(?:(?<endSheet>'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_\.]*)!)?(?<endColAbs>\$?)(?<endCol>[A-Za-z]{1,3})(?<endRowAbs>\$?)(?<endRow>\d{1,7})(?=[:),+\-*/^&=<>%# \t\r\n]|$)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant,
            StructuralReferenceRegexTimeout);

        internal static readonly Regex RowRangeReferenceRegex = new Regex(
            @"(?<![\p{L}\p{N}_\.\]:!\\])(?:(?<sheet>'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_\.]*)!)?(?<startRowAbs>\$?)(?<startRow>\d{1,7}):(?:(?<endSheet>'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_\.]*)!)?(?<endRowAbs>\$?)(?<endRow>\d{1,7})(?=[:),+\-*/^&=<>%# \t\r\n]|$)",
            RegexOptions.Compiled | RegexOptions.CultureInvariant,
            StructuralReferenceRegexTimeout);

        internal static string RewriteOutsideStrings(
            string formula,
            Func<string, string> rewriteSegment) {
            if (formula == null) throw new ArgumentNullException(nameof(formula));
            if (rewriteSegment == null) throw new ArgumentNullException(nameof(rewriteSegment));

            var builder = new StringBuilder(formula.Length);
            int index = 0;
            while (index < formula.Length) {
                int segmentStart = index;
                bool insideSingleQuotedQualifier = false;
                while (index < formula.Length) {
                    char character = formula[index];
                    if (character == '\'') {
                        if (insideSingleQuotedQualifier
                            && index + 1 < formula.Length
                            && formula[index + 1] == '\'') {
                            index += 2;
                            continue;
                        }

                        insideSingleQuotedQualifier = !insideSingleQuotedQualifier;
                        index++;
                        continue;
                    }

                    if (character == '"' && !insideSingleQuotedQualifier) {
                        break;
                    }

                    index++;
                }

                if (index >= formula.Length) {
                    builder.Append(rewriteSegment(formula.Substring(segmentStart)));
                    break;
                }

                if (index > segmentStart) {
                    builder.Append(rewriteSegment(formula.Substring(segmentStart, index - segmentStart)));
                }

                int literalStart = index;
                index++;
                while (index < formula.Length) {
                    if (formula[index] == '"') {
                        if (index + 1 < formula.Length && formula[index + 1] == '"') {
                            index += 2;
                            continue;
                        }

                        index++;
                        break;
                    }

                    index++;
                }

                builder.Append(formula, literalStart, index - literalStart);
            }

            return builder.ToString();
        }
    }
}
