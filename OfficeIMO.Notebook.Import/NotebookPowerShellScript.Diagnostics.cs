using System.Globalization;
using System.Text.Json;

namespace OfficeIMO.Notebook.Import;

public static partial class NotebookPowerShellScript {
    private static bool CanExport(NotebookCell cell, string? defaultLanguage, int index, List<OfficeDiagnostic> diagnostics) {
        if (cell.Kind == NotebookCellKind.Markdown) return true;
        if (cell.Kind == NotebookCellKind.Code) {
            if (IsPowerShell(cell.Language ?? defaultLanguage)) return true;
            AddDiagnostic(diagnostics, "notebook-powershell-language-omitted",
                "Code that was not identified as PowerShell was omitted from the script.", Location(index));
        } else {
            AddDiagnostic(diagnostics, "notebook-powershell-cell-omitted",
                "The PowerShell script profile omits raw and unsupported cell kinds.", Location(index));
        }
        return false;
    }

    private static void ReportCellLosses(NotebookCell cell, int index, List<OfficeDiagnostic> diagnostics) {
        string location = Location(index);
        if (cell.Outputs.Count != 0) {
            AddDiagnostic(diagnostics, "notebook-powershell-outputs-omitted",
                "Saved cell outputs were omitted from the script.", location + ".outputs");
        }
        if (cell.ExecutionCount.HasValue) {
            AddDiagnostic(diagnostics, "notebook-powershell-execution-count-omitted",
                "The saved execution count was omitted from the script.", location + ".execution_count");
        }
        if (cell.Attachments.Count != 0) {
            AddDiagnostic(diagnostics, "notebook-powershell-attachments-omitted",
                "Notebook-local attachments were omitted from the script.", location + ".attachments");
        }

        bool hasMetadata = cell.Metadata.Any(entry => !(cell.Kind == NotebookCellKind.Code && entry.Key == "language" &&
            entry.Value.ValueKind == JsonValueKind.String && IsPowerShell(entry.Value.GetString())));
        if (hasMetadata || (cell.Original.ValueKind == JsonValueKind.Object && cell.Original.EnumerateObject().Any(static property =>
            property.Name is not ("cell_type" or "id" or "source" or "metadata" or "outputs" or "execution_count" or "attachments")))) {
            AddDiagnostic(diagnostics, "notebook-powershell-cell-metadata-omitted",
                "Cell metadata and extension properties outside the PowerShell script profile were omitted.", location + ".metadata");
        }
    }

    private static bool IsNotebookLanguageMetadata(KeyValuePair<string, JsonElement> entry) {
        if (entry.Key != "language_info" || entry.Value.ValueKind != JsonValueKind.Object) return false;
        JsonElement.ObjectEnumerator properties = entry.Value.EnumerateObject();
        if (!properties.MoveNext()) return false;
        JsonProperty property = properties.Current;
        return property.NameEquals("name") && property.Value.ValueKind == JsonValueKind.String &&
            IsPowerShell(property.Value.GetString()) && !properties.MoveNext();
    }

    private static bool IsPowerShell(string? language) =>
        string.Equals(language, Language, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(language, "pwsh", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(language, "ps1", StringComparison.OrdinalIgnoreCase);

    private static string Location(int index) => "cells[" + index.ToString(CultureInfo.InvariantCulture) + "]";

    private static void AddDiagnostic(List<OfficeDiagnostic> diagnostics, string code, string message, string? location) {
        if (diagnostics.Count < MaximumDiagnosticDetails) {
            diagnostics.Add(new OfficeDiagnostic(code, OfficeDiagnosticSeverity.Warning, message, location, TargetFormat,
                OfficeConversionLossKind.Omission));
        } else if (diagnostics.Count == MaximumDiagnosticDetails) {
            AddDiagnosticLimit(diagnostics);
        }
    }

    private static void AddDiagnosticLimit(List<OfficeDiagnostic> diagnostics) =>
        diagnostics.Add(new OfficeDiagnostic("notebook-powershell-diagnostic-limit", OfficeDiagnosticSeverity.Warning,
            "Additional PowerShell script conversion findings were suppressed after reaching the report limit.",
            format: TargetFormat, lossKind: OfficeConversionLossKind.Unassessed));
}
