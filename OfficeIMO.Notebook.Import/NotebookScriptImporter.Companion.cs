using System.Globalization;
using System.Text.Json;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookScriptImporter {
    private static void AttachCompanion(NotebookImportContext context, NotebookDocument companion, IReadOnlyList<string?> explicitIds, string scriptLanguage) {
        IReadOnlyList<NotebookCell> scriptCells = context.Document.Cells;
        IReadOnlyList<NotebookCell> pairedCells = companion.Cells;
        if (scriptCells.Count != pairedCells.Count) {
            ReportPairMismatch(context, "The script and its Jupyter companion contain different numbers of cells.");
            return;
        }

        var seenCode = new Dictionary<string, (bool HasSavedResult, bool IsIdentified)>(StringComparer.Ordinal);
        for (int i = 0; i < scriptCells.Count; i++) {
            context.CancellationToken.ThrowIfCancellationRequested();
            NotebookCell script = scriptCells[i];
            NotebookCell paired = pairedCells[i];
            string? explicitId = explicitIds[i];
            string? declaredPairedId = paired.Original.ValueKind == JsonValueKind.Undefined
                ? paired.Id : ReadStringProperty(paired.Original, "id");
            bool isIdentified = explicitId != null && string.Equals(explicitId, script.Id, StringComparison.Ordinal)
                && string.Equals(explicitId, paired.Id, StringComparison.Ordinal)
                && string.Equals(explicitId, declaredPairedId, StringComparison.Ordinal);
            string scriptSource = script.Source;
            string pairedSource = NormalizeLineBreaks(paired.Source);
            if (script.Kind != paired.Kind || !string.Equals(scriptSource, pairedSource, StringComparison.Ordinal)
                || (explicitId != null && !isIdentified)) {
                ReportPairMismatch(context, "The script and its Jupyter companion differ in cell source, type, or explicit ID.", i);
                return;
            }
            if (script.Kind != NotebookCellKind.Code) continue;

            string? pairedLanguage = paired.Language ?? companion.Language;
            string? expectedLanguage = paired.Language == null ? scriptLanguage : script.Language;
            if (!string.IsNullOrWhiteSpace(pairedLanguage) && !SameLanguage(expectedLanguage, pairedLanguage!)) {
                ReportPairMismatch(context, "The script and its Jupyter companion declare different code languages.", i);
                return;
            }

            if (paired.ExecutionCount < 0) {
                ReportPairMismatch(context, "The Jupyter companion has an invalid execution count.", i);
                return;
            }
            bool hasSavedResult = paired.Outputs.Count > 0 || paired.ExecutionCount != null;
            if (seenCode.TryGetValue(scriptSource, out var previous)) {
                if ((hasSavedResult || previous.HasSavedResult) && (!isIdentified || !previous.IsIdentified)) {
                    ReportPairMismatch(context, "Repeated code with saved results requires matching explicit IDs in each occurrence.", i);
                    return;
                }
                seenCode[scriptSource] = (hasSavedResult || previous.HasSavedResult, isIdentified && previous.IsIdentified);
            } else {
                seenCode.Add(scriptSource, (hasSavedResult, isIdentified));
            }
        }

        for (int i = 0; i < scriptCells.Count; i++) {
            context.CancellationToken.ThrowIfCancellationRequested();
            NotebookCell target = scriptCells[i];
            if (target.Kind != NotebookCellKind.Code) continue;
            NotebookCell saved = pairedCells[i];
            target.ExecutionCount = saved.ExecutionCount;
            foreach (NotebookOutput output in saved.Outputs) context.AddOutput(target, output);
        }
    }

    private static string NormalizeLineBreaks(string source) => source.IndexOf('\r') < 0
        ? source : source.Replace("\r\n", "\n").Replace('\r', '\n');

    private static bool SameLanguage(string? script, string paired) {
        if (script == null || string.Equals(script.Trim(), paired.Trim(), StringComparison.OrdinalIgnoreCase)) return true;
        return LanguageAlias(script) == LanguageAlias(paired);
    }

    private static string LanguageAlias(string value) => value.Trim().ToLowerInvariant() switch {
        "py" or "python2" or "python3" or "ipython" => "python",
        "ir" => "r",
        "jl" => "julia",
        var language => language
    };

    private static void ReportPairMismatch(NotebookImportContext context, string message, int? index = null) =>
        context.Report("notebook-companion-mismatch", message + " Saved results were omitted.",
            index == null ? null : "cells[" + index.Value.ToString(CultureInfo.InvariantCulture) + "]");
}
