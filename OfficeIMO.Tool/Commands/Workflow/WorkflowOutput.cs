using System.Globalization;
using System.Text;
using System.Text.Json;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

/// <summary>JSON and text results for export-pages, assemble and print-plan.</summary>
internal static class WorkflowOutput {
    internal const string ExportPagesSchema = "officeimo.workflow.export-pages.v1";
    internal const string AssembleSchema = "officeimo.workflow.assemble.v1";
    internal const string PrintPlanSchema = "officeimo.workflow.print-plan.v1";

    internal static async Task WriteExportPagesAsync(TextWriter output, PdfPageImageExportResult result, string inputPath, bool json) {
        if (!json) {
            if (!result.Succeeded) return;
            await output.WriteLineAsync(result.OutputDirectory).ConfigureAwait(false);
            await output.WriteLineAsync(Count(result.Files.Count) + " page image(s), " + Count(result.OutputBytes) + " bytes").ConfigureAwait(false);
            return;
        }
        await WriteJsonAsync(output, writer => {
            WriteStatus(writer, ExportPagesSchema, result.Status, result.FailureKind, result.Summary);
            writer.WriteString("inputPath", inputPath);
            writer.WriteString("outputDirectory", result.OutputDirectory);
            writer.WriteNumber("fileCount", result.Files.Count);
            writer.WriteNumber("inputBytes", result.InputBytes);
            writer.WriteNumber("outputBytes", result.OutputBytes);
            writer.WriteStartArray("files");
            foreach (PdfPageImageFile file in result.Files) {
                writer.WriteStartObject();
                writer.WriteNumber("pageNumber", file.PageNumber);
                writer.WriteString("path", file.Path);
                writer.WriteString("format", file.Format.ToString().ToLowerInvariant());
                writer.WriteNumber("width", file.Width);
                writer.WriteNumber("height", file.Height);
                writer.WriteNumber("sizeBytes", file.SizeBytes);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
            WriteDiagnostics(writer, result.Diagnostics);
        }).ConfigureAwait(false);
    }

    internal static async Task WriteAssembleAsync(TextWriter output, PdfAssemblyResult result, bool json) {
        if (!json) {
            if (!result.Succeeded) return;
            await output.WriteLineAsync(result.OutputPath).ConfigureAwait(false);
            await output.WriteLineAsync(Count(result.PageCount) + " page(s) from " + Count(result.SourceCount) + " normalized source(s)").ConfigureAwait(false);
            return;
        }
        await WriteJsonAsync(output, writer => {
            WriteStatus(writer, AssembleSchema, result.Status, result.FailureKind, result.Summary);
            writer.WriteString("outputPath", result.OutputPath);
            writer.WriteNumber("sourceCount", result.SourceCount);
            writer.WriteNumber("pageCount", result.PageCount);
            writer.WriteNumber("inputBytes", result.InputBytes);
            writer.WriteNumber("outputBytes", result.OutputBytes);
            WriteDiagnostics(writer, result.Diagnostics);
        }).ConfigureAwait(false);
    }

    internal static async Task WritePrintPlanAsync(TextWriter output, PdfPrintPlan plan, string inputPath, bool json) {
        if (!json) {
            await output.WriteLineAsync("Source pages: " + Count(plan.SourcePageCount)).ConfigureAwait(false);
            await output.WriteLineAsync("Selected pages: " + string.Join(',', plan.SelectedPages)).ConfigureAwait(false);
            await output.WriteLineAsync("Sheets: " + Count(plan.Sheets.Count)).ConfigureAwait(false);
            await output.WriteLineAsync("Color: " + plan.ColorMode).ConfigureAwait(false);
            foreach (PdfPrintSheet sheet in plan.Sheets) {
                await output.WriteLineAsync(
                    "Sheet " + Count(sheet.SheetNumber) + ": " +
                    sheet.PaperSize.Width.ToPoints().ToString("0.##", CultureInfo.InvariantCulture) + " x " +
                    sheet.PaperSize.Height.ToPoints().ToString("0.##", CultureInfo.InvariantCulture) + " pt; pages " +
                    string.Join(',', sheet.Placements.Select(static placement => placement.PageNumber))).ConfigureAwait(false);
            }
            return;
        }
        await WriteJsonAsync(output, writer => {
            writer.WriteString("schema", PrintPlanSchema);
            writer.WriteString("inputPath", inputPath);
            writer.WriteNumber("sourcePageCount", plan.SourcePageCount);
            writer.WriteStartArray("selectedPages");
            foreach (int page in plan.SelectedPages) writer.WriteNumberValue(page);
            writer.WriteEndArray();
            writer.WriteNumber("sheetCount", plan.Sheets.Count);
            writer.WriteString("colorMode", plan.ColorMode.ToString().ToLowerInvariant());
            writer.WriteStartArray("sheets");
            foreach (PdfPrintSheet sheet in plan.Sheets) {
                writer.WriteStartObject();
                writer.WriteNumber("sheetNumber", sheet.SheetNumber);
                writer.WriteNumber("paperWidth", Points(sheet.PaperSize.Width.ToPoints()));
                writer.WriteNumber("paperHeight", Points(sheet.PaperSize.Height.ToPoints()));
                writer.WriteStartArray("placements");
                foreach (PdfPrintPlacement placement in sheet.Placements) {
                    writer.WriteStartObject();
                    writer.WriteNumber("pageNumber", placement.PageNumber);
                    writer.WriteNumber("x", Points(placement.X));
                    writer.WriteNumber("y", Points(placement.Y));
                    writer.WriteNumber("width", Points(placement.Width));
                    writer.WriteNumber("height", Points(placement.Height));
                    writer.WriteNumber("scale", Math.Round(placement.Scale, 4));
                    writer.WriteBoolean("clipped", placement.IsClipped);
                    writer.WriteEndObject();
                }
                writer.WriteEndArray();
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }).ConfigureAwait(false);
    }

    private static void WriteStatus(Utf8JsonWriter writer, string schema, OfficeWorkflowStatus status, OfficeWorkflowFailureKind failureKind, string summary) {
        writer.WriteString("schema", schema);
        writer.WriteString("status", status.ToString());
        writer.WriteString("failureKind", failureKind.ToString());
        writer.WriteBoolean("succeeded", status == OfficeWorkflowStatus.Completed);
        writer.WriteString("summary", summary);
    }

    private static void WriteDiagnostics(Utf8JsonWriter writer, IReadOnlyList<OfficeWorkflowDiagnostic> diagnostics) {
        writer.WriteStartArray("diagnostics");
        foreach (OfficeWorkflowDiagnostic diagnostic in diagnostics) {
            writer.WriteStartObject();
            writer.WriteString("severity", diagnostic.Severity.ToString());
            writer.WriteString("code", diagnostic.Code);
            writer.WriteString("message", diagnostic.Message);
            writer.WriteEndObject();
        }
        writer.WriteEndArray();
    }

    private static async Task WriteJsonAsync(TextWriter output, Action<Utf8JsonWriter> write) {
        using var bytes = new MemoryStream();
        using (var writer = new Utf8JsonWriter(bytes, new JsonWriterOptions { Indented = true })) {
            writer.WriteStartObject();
            write(writer);
            writer.WriteEndObject();
        }
        await output.WriteLineAsync(Encoding.UTF8.GetString(bytes.ToArray())).ConfigureAwait(false);
    }

    private static double Points(double value) => Math.Round(value, 2);

    private static string Count(long value) => value.ToString(CultureInfo.InvariantCulture);
}
