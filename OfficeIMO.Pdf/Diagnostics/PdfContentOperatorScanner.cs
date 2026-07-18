namespace OfficeIMO.Pdf;

internal static class PdfContentOperatorScanner {
    internal static void AppendOperators(string content, List<string> destination, int maximum, ref bool truncated) {
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

                if (destination.Count >= maximum) {
                    return false;
                }

                destination.Add(operation.Name);
                return true;
            });
        if (!complete) {
            truncated = true;
        }
    }
}
