using System.Globalization;
using System.Text;
using System.Text.Json;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Workflow;

/// <summary>JSON and text results for the workflow commands.</summary>
internal static class WorkflowOutput {
    internal const string ExportPagesSchema = "officeimo.workflow.export-pages.v1";
    internal const string AssembleSchema = "officeimo.workflow.assemble.v1";
    internal const string PrintPlanSchema = "officeimo.workflow.print-plan.v1";
    internal const string CompareSchema = "officeimo.workflow.compare.v1";
    internal const string OptimizeImagesSchema = "officeimo.workflow.optimize-images.v1";
    internal const string PrintersSchema = "officeimo.workflow.printers.v1";
    internal const string PaperSourcesSchema = "officeimo.workflow.paper-sources.v1";
    internal const string PrintSchema = "officeimo.workflow.print.v1";

    /// <summary>Reads a --format value; <paramref name="otherUse"/> names the option for a non-output value, if any.</summary>
    internal static bool ParseJson(string value, string? otherUse = null) => value.ToLowerInvariant() switch {
        "json" => true,
        "text" => false,
        _ => throw new WorkflowUsageException(otherUse == null
            ? "--format must be json or text."
            : "--format selects json or text output. Use " + otherUse + " " + value + " instead.")
    };

    internal static async Task WriteCompareAsync(TextWriter output, OfficeWorkflowResult result, string inputPath, string comparisonPath, bool json) {
        if (!json) {
            if (!result.Succeeded) return;
            await output.WriteLineAsync(result.OutputPath).ConfigureAwait(false);
            await output.WriteLineAsync(result.Summary).ConfigureAwait(false);
            return;
        }
        await WriteJsonAsync(output, writer => {
            WriteStatus(writer, CompareSchema, result.Status, result.FailureKind, result.Summary);
            writer.WriteString("inputPath", inputPath);
            writer.WriteString("comparisonPath", comparisonPath);
            writer.WriteString("outputPath", result.OutputPath);
            writer.WriteNumber("inputBytes", result.InputBytes);
            writer.WriteNumber("outputBytes", result.OutputBytes);
            writer.WriteString("outputSha256", result.OutputSha256);
            WriteDiagnostics(writer, result.Diagnostics);
        }).ConfigureAwait(false);
    }

    internal static async Task WriteOptimizeImagesAsync(TextWriter output, TextWriter error, IReadOnlyList<OfficeWorkflowResult> results, bool analyze, bool json) {
        if (!json) {
            foreach (OfficeWorkflowResult result in results) {
                await output.WriteLineAsync((result.OutputPath ?? result.RequestId) + ": " + result.Summary).ConfigureAwait(false);
                foreach (OfficeWorkflowDiagnostic diagnostic in result.Diagnostics)
                    await (result.Succeeded ? output : error).WriteLineAsync(diagnostic.Code + ": " + diagnostic.Message).ConfigureAwait(false);
            }
            return;
        }
        foreach (OfficeWorkflowResult result in results.Where(static result => !result.Succeeded))
            foreach (OfficeWorkflowDiagnostic diagnostic in result.Diagnostics)
                await error.WriteLineAsync(result.RequestId + ": " + diagnostic.Code + ": " + diagnostic.Message).ConfigureAwait(false);
        await WriteJsonAsync(output, writer => {
            writer.WriteString("schema", OptimizeImagesSchema);
            writer.WriteString("mode", analyze ? "analyze" : "optimize");
            writer.WriteBoolean("succeeded", results.All(static result => result.Succeeded));
            writer.WriteNumber("fileCount", results.Count);
            writer.WriteStartArray("files");
            foreach (OfficeWorkflowResult result in results) {
                writer.WriteStartObject();
                writer.WriteString("inputPath", result.RequestId);
                writer.WriteString("outputPath", result.OutputPath);
                writer.WriteString("status", result.Status.ToString());
                writer.WriteString("failureKind", result.FailureKind.ToString());
                writer.WriteBoolean("succeeded", result.Succeeded);
                writer.WriteString("summary", result.Summary);
                writer.WriteNumber("inputBytes", result.InputBytes);
                writer.WriteNumber("outputBytes", result.OutputBytes);
                WriteDiagnostics(writer, result.Diagnostics);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }).ConfigureAwait(false);
    }

    internal static async Task WritePrintersAsync(TextWriter output, IReadOnlyList<PdfPrinterInfo> printers, bool json) {
        if (!json) {
            foreach (PdfPrinterInfo queue in printers)
                await output.WriteLineAsync(queue.Name + (queue.IsDefault ? " (default)" : "") + (queue.RequiresOutputFile ? " (file printer)" : "")).ConfigureAwait(false);
            return;
        }
        await WriteJsonAsync(output, writer => {
            writer.WriteString("schema", PrintersSchema);
            writer.WriteStartArray("printers");
            foreach (PdfPrinterInfo queue in printers) {
                writer.WriteStartObject();
                writer.WriteString("name", queue.Name);
                writer.WriteBoolean("isDefault", queue.IsDefault);
                writer.WriteBoolean("requiresOutputFile", queue.RequiresOutputFile);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }).ConfigureAwait(false);
    }

    internal static async Task WritePaperSourcesAsync(TextWriter output, string printer, IReadOnlyList<PdfPaperSourceInfo> sources, bool json) {
        if (!json) {
            foreach (PdfPaperSourceInfo source in sources)
                await output.WriteLineAsync(source.Id + ": " + source.Name).ConfigureAwait(false);
            return;
        }
        await WriteJsonAsync(output, writer => {
            writer.WriteString("schema", PaperSourcesSchema);
            writer.WriteString("printer", printer);
            writer.WriteStartArray("paperSources");
            foreach (PdfPaperSourceInfo source in sources) {
                writer.WriteStartObject();
                writer.WriteString("id", source.Id);
                writer.WriteString("name", source.Name);
                writer.WriteEndObject();
            }
            writer.WriteEndArray();
        }).ConfigureAwait(false);
    }

    internal static async Task WritePrintAsync(TextWriter output, PdfPrintSubmission receipt, string inputPath, bool json) {
        if (!json) {
            await output.WriteLineAsync("Queue accepted: " + receipt.PrinterName + "; job " + (receipt.JobId ?? "unknown") + ". Physical delivery is unconfirmed.").ConfigureAwait(false);
            return;
        }
        await WriteJsonAsync(output, writer => {
            writer.WriteString("schema", PrintSchema);
            writer.WriteString("inputPath", inputPath);
            writer.WriteString("printer", receipt.PrinterName);
            writer.WriteString("jobId", receipt.JobId);
            writer.WriteNumber("sheetCount", receipt.SheetCount);
            writer.WriteNumber("copies", receipt.Copies);
            writer.WriteString("outputFilePath", receipt.OutputFilePath);
        }).ConfigureAwait(false);
    }

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
            if (diagnostic.Details.Count != 0) {
                writer.WriteStartObject("details");
                foreach (KeyValuePair<string, string> detail in diagnostic.Details.OrderBy(static pair => pair.Key, StringComparer.Ordinal))
                    writer.WriteString(detail.Key, detail.Value);
                writer.WriteEndObject();
            }
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
