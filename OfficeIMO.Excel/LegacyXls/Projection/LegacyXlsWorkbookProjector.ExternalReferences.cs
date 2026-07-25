using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Runtime.CompilerServices;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static partial class LegacyXlsWorkbookProjector {
        private static readonly ConditionalWeakTable<LegacyXlsWorkbook, ExternalWorkbookReferenceMatcher>
            ExternalWorkbookReferenceMatchers = new();

        private sealed class ExternalWorkbookReferenceMatcher {
            private readonly HashSet<string> _fileNames = new(StringComparer.OrdinalIgnoreCase);

            internal ExternalWorkbookReferenceMatcher(IReadOnlyList<LegacyXlsExternalReference> references) {
                foreach (LegacyXlsExternalReference reference in references) {
                    if (reference.Kind != LegacyXlsExternalReferenceKind.ExternalWorkbook ||
                        string.IsNullOrWhiteSpace(reference.Target)) {
                        continue;
                    }

                    string fileName = BiffFormulaReferenceFormatter.NormalizeExternalWorkbookTarget(reference.Target);
                    if (fileName.Length > 0) _fileNames.Add(fileName);
                }
            }

            internal bool ReferencesExternalWorkbook(string formulaText) {
                if (_fileNames.Count == 0 || string.IsNullOrEmpty(formulaText)) return false;

                bool inStringLiteral = false;
                for (int index = 0; index < formulaText.Length; index++) {
                    char current = formulaText[index];
                    if (current == '"') {
                        if (inStringLiteral && index + 1 < formulaText.Length && formulaText[index + 1] == '"') {
                            index++;
                        } else {
                            inStringLiteral = !inStringLiteral;
                        }
                        continue;
                    }
                    if (inStringLiteral) continue;

                    if (current == '\'') {
                        int end = FindQuotedQualifierEnd(formulaText, index + 1);
                        if (end < 0) return false;
                        if (end + 1 < formulaText.Length && formulaText[end + 1] == '!' &&
                            QualifierMatches(formulaText.Substring(index + 1, end - index - 1).Replace("''", "'"))) {
                            return true;
                        }
                        index = end;
                        continue;
                    }

                    if (current == '[') {
                        int qualifierEnd = formulaText.IndexOf('!', index + 1);
                        int end = qualifierEnd < 0
                            ? -1
                            : formulaText.LastIndexOf(']', qualifierEnd - 1, qualifierEnd - index - 1);
                        if (end > index + 1 && _fileNames.Contains(formulaText.Substring(index + 1, end - index - 1))) {
                            return true;
                        }
                    } else if (current == '!') {
                        int start = index - 1;
                        while (start >= 0 && !IsQualifierBoundary(formulaText[start])) start--;
                        start++;
                        if (start < index && _fileNames.Contains(formulaText.Substring(start, index - start))) {
                            return true;
                        }
                    }
                }

                return false;
            }

            private bool QualifierMatches(string qualifier) {
                int open = qualifier.IndexOf('[');
                int close = open < 0 ? -1 : qualifier.LastIndexOf(']');
                return open >= 0 && close > open + 1
                    ? _fileNames.Contains(qualifier.Substring(open + 1, close - open - 1))
                    : _fileNames.Contains(qualifier);
            }

            private static int FindQuotedQualifierEnd(string formulaText, int start) {
                for (int index = start; index < formulaText.Length; index++) {
                    if (formulaText[index] != '\'') continue;
                    if (index + 1 < formulaText.Length && formulaText[index + 1] == '\'') {
                        index++;
                        continue;
                    }
                    return index;
                }
                return -1;
            }

            private static bool IsQualifierBoundary(char value) =>
                char.IsWhiteSpace(value) || "+-*/^&=<>(),;:{}".IndexOf(value) >= 0;
        }
    }
}
