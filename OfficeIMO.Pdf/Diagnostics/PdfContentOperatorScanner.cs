namespace OfficeIMO.Pdf;

internal static class PdfContentOperatorScanner {
    internal static void AppendOperators(
        string content,
        List<string> destination,
        int maximum,
        ref bool truncated,
        int maxNestingDepth = PdfReadLimits.DefaultMaxContentNestingDepth,
        int maxOperands = PdfReadLimits.DefaultMaxContentOperands) {
        if (destination.Count >= maximum) {
            truncated = true;
            return;
        }

        bool complete = PdfContentStreamInterpreter.InterpretUntil(
            content,
            int.MaxValue,
            operation => {
                if (!PdfContentOperators.IsStandard(operation.Name)) {
                    return true;
                }

                if (!Append(operation.Name)) return false;
                if (string.Equals(operation.Name, "BI", StringComparison.Ordinal) &&
                    operation.InlineImage is not null) {
                    if (!Append("ID")) return false;
                    if (!Append("EI")) return false;
                }
                return true;

                bool Append(string name) {
                    if (destination.Count >= maximum) {
                        return false;
                    }

                    destination.Add(name);
                    return true;
                }
            },
            maxNestingDepth: maxNestingDepth,
            maxOperands: maxOperands);
        if (!complete) {
            truncated = true;
        }
    }
}
