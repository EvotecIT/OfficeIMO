using System;
using System.Collections.Generic;
using System.Text;

namespace OfficeIMO.Excel {
    internal sealed class ExcelFormulaReferenceCandidate {
        internal ExcelFormulaReferenceCandidate(int index, int length, string text, ExcelReference reference, bool hasSpill) {
            Index = index;
            Length = length;
            Text = text;
            Reference = reference;
            HasSpill = hasSpill;
        }

        internal int Index { get; }
        internal int Length { get; }
        internal string Text { get; }
        internal ExcelReference Reference { get; }
        internal bool HasSpill { get; }
    }

    /// <summary>
    /// Owns the bounded lexical grammar used to discover A1 references in formula-bearing workbook metadata.
    /// </summary>
    /// <remarks>
    /// Consumers remain responsible for deciding whether a discovered reference belongs to the edited sheet
    /// and how it should move. Keeping discovery here prevents shared formulas, structural edits, names,
    /// charts, validations, and template operations from drifting into incompatible reference grammars.
    /// </remarks>
    internal static class ExcelFormulaReferenceRewriter {
        /// <summary>
        /// Reads one complete A1 reference at an exact formula cursor without regular expressions.
        /// Quoted sheet names and bracketed external-workbook qualifiers are scanned as opaque lexical regions,
        /// then the candidate is validated by the canonical <see cref="ExcelReference"/> parser.
        /// </summary>
        internal static bool TryReadReferenceAt(
            string formula,
            int start,
            out ExcelFormulaReferenceCandidate? candidate) {
            candidate = null;
            if (formula == null) throw new ArgumentNullException(nameof(formula));
            if (start < 0 || start >= formula.Length) return false;
            if (start > 0 && IsReferenceIdentifierPart(formula[start - 1])) return false;
            if (!CanStartReference(formula, start)) return false;

            int index = start;
            bool quotedQualifier = false;
            int bracketDepth = 0;
            while (index < formula.Length) {
                char value = formula[index];
                if (quotedQualifier) {
                    if (value == '\'' && index + 1 < formula.Length && formula[index + 1] == '\'') {
                        index += 2;
                        continue;
                    }
                    if (value == '\'') quotedQualifier = false;
                    index++;
                    continue;
                }
                if (value == '\'') {
                    quotedQualifier = true;
                    index++;
                    continue;
                }
                if (value == '[') {
                    bracketDepth++;
                    index++;
                    continue;
                }
                if (value == ']' && bracketDepth > 0) {
                    bracketDepth--;
                    index++;
                    continue;
                }
                if (bracketDepth == 0 && IsFormulaReferenceTerminator(value)) break;
                index++;
            }

            if (quotedQualifier || bracketDepth != 0 || index == start) return false;
            int candidateEnd = index;
            bool hasSpill = formula[candidateEnd - 1] == '#';
            int referenceEnd = candidateEnd - (hasSpill ? 1 : 0);
            // Check the repeated-qualified Excel spelling first. ExcelReference.TryParse is intentionally
            // forgiving and can otherwise accept only the final endpoint from text such as
            // Sheet1!A1:Sheet1!C3, silently losing the source span and the first endpoint.
            if (!TryParseRepeatedQualifiedRange(formula, start, referenceEnd, out ExcelReference? reference)
                && !TryParseCandidate(formula, start, referenceEnd, out reference)) {
                // Excel accepts a repeated qualifier on the second endpoint (Sheet!A1:Sheet!C3),
                // while the canonical reference parser intentionally models one qualifier for a range.
                // Return the first cell here so the syntax tree can combine the two qualified endpoints.
                int rangeSeparator = FindLastRangeSeparator(formula, start, referenceEnd);
                if (rangeSeparator <= start
                    || !TryParseCandidate(formula, start, rangeSeparator, out reference)) {
                    return false;
                }
                candidateEnd = rangeSeparator;
                referenceEnd = rangeSeparator;
                hasSpill = false;
            }
            if (index < formula.Length && IsReferenceIdentifierPart(formula[index])) return false;
            candidate = new ExcelFormulaReferenceCandidate(
                start,
                candidateEnd - start,
                formula.Substring(start, candidateEnd - start),
                reference!,
                hasSpill);
            return true;
        }

        private static bool TryParseCandidate(
            string formula,
            int start,
            int end,
            out ExcelReference? reference) {
            reference = null;
            if (end <= start) return false;
            return ExcelReference.TryParse(formula.Substring(start, end - start), out reference);
        }

        private static int FindLastRangeSeparator(string formula, int start, int end) {
            bool quotedQualifier = false;
            int bracketDepth = 0;
            for (int index = end - 1; index >= start; index--) {
                char value = formula[index];
                if (value == ']' && !quotedQualifier) { bracketDepth++; continue; }
                if (value == '[' && !quotedQualifier && bracketDepth > 0) { bracketDepth--; continue; }
                if (bracketDepth > 0) continue;
                if (value == '\'') {
                    if (index > start && formula[index - 1] == '\'') { index--; continue; }
                    quotedQualifier = !quotedQualifier;
                    continue;
                }
                if (!quotedQualifier && value == ':') return index;
            }
            return -1;
        }

        private static bool TryParseRepeatedQualifiedRange(
            string formula,
            int start,
            int end,
            out ExcelReference? reference) {
            reference = null;
            int separator = FindLastRangeSeparator(formula, start, end);
            if (separator <= start || separator + 1 >= end) return false;
            string first = formula.Substring(start, separator - start);
            string second = formula.Substring(separator + 1, end - separator - 1);
            if (!TrySplitQualifiedEndpoint(first, out string firstQualifier, out string firstEndpoint)
                || !TrySplitQualifiedEndpoint(second, out string secondQualifier, out string secondEndpoint)
                || !string.Equals(
                    ExcelReference.NormalizeQualifierForComparison(firstQualifier),
                    ExcelReference.NormalizeQualifierForComparison(secondQualifier),
                    StringComparison.OrdinalIgnoreCase)) return false;
            return ExcelReference.TryParse(
                firstQualifier + "!" + firstEndpoint + ":" + secondEndpoint,
                out reference);
        }

        private static bool TrySplitQualifiedEndpoint(
            string value,
            out string qualifier,
            out string endpoint) {
            qualifier = string.Empty;
            endpoint = string.Empty;
            bool quoted = false;
            int separator = -1;
            for (int index = 0; index < value.Length; index++) {
                if (value[index] == '\'') {
                    if (quoted && index + 1 < value.Length && value[index + 1] == '\'') { index++; continue; }
                    quoted = !quoted;
                } else if (value[index] == '!' && !quoted) {
                    separator = index;
                }
            }
            if (quoted || separator <= 0 || separator + 1 >= value.Length) return false;
            qualifier = value.Substring(0, separator);
            endpoint = value.Substring(separator + 1);
            return true;
        }

        private static bool IsFormulaReferenceTerminator(char value) =>
            char.IsWhiteSpace(value) || value == '(' || value == ')' || value == '{' || value == '}' ||
            value == ',' || value == ';' || value == '~' || value == '+' || value == '-' || value == '*' ||
            value == '/' || value == '^' || value == '&' || value == '=' || value == '<' || value == '>' ||
            value == '%' || value == '"';

        private static bool IsReferenceIdentifierPart(char value) =>
            char.IsLetterOrDigit(value) || value == '_' || value == '.' || value == '\\';

        private static bool CanStartReference(string formula, int start) {
            char value = formula[start];
            if (value == '\'' || value == '[' || value == '_' || char.IsLetterOrDigit(value)) return true;
            if (value != '$' || start + 1 >= formula.Length) return false;
            char next = formula[start + 1];
            return char.IsLetterOrDigit(next);
        }

        /// <summary>Returns syntax-validated references with their exact source spans.</summary>
        internal static IReadOnlyList<ExcelFormulaReferenceCandidate> FindReferences(string formula) {
            if (formula == null) throw new ArgumentNullException(nameof(formula));
            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(formula);
            var references = new List<ExcelFormulaReferenceCandidate>();
            int index = 0;
            foreach (ExcelFormulaSyntaxNode node in tree.Nodes) {
                if (node is ExcelFormulaReferenceSyntax reference) {
                    references.Add(new ExcelFormulaReferenceCandidate(
                        index,
                        node.Text.Length,
                        node.Text,
                        reference.Reference,
                        node.Text.EndsWith("#", StringComparison.Ordinal)));
                }
                index += node.Text.Length;
            }
            return references;
        }

        /// <summary>Rewrites complete syntax-validated references while preserving every other token verbatim.</summary>
        internal static string RewriteReferences(
            string formula,
            Func<ExcelFormulaReferenceCandidate, string> rewriter) {
            if (formula == null) throw new ArgumentNullException(nameof(formula));
            if (rewriter == null) throw new ArgumentNullException(nameof(rewriter));
            IReadOnlyList<ExcelFormulaReferenceCandidate> references = FindReferences(formula);
            if (references.Count == 0) return formula;
            var builder = new StringBuilder(formula.Length);
            int cursor = 0;
            foreach (ExcelFormulaReferenceCandidate reference in references) {
                builder.Append(formula, cursor, reference.Index - cursor);
                builder.Append(rewriter(reference));
                cursor = reference.Index + reference.Length;
            }
            builder.Append(formula, cursor, formula.Length - cursor);
            return builder.ToString();
        }

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
