using OfficeIMO.Ocr;
using OfficeIMO.Tool.Agent;
using OfficeIMO.Tool.Commands.Workflow;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Pdf;

internal static class PdfWorkflowCommand {
    internal const string Usage = """
OfficeIMO.Tool - PDF workflows

Usage:
  officeimo pdf extract <input.pdf> --pages <selection> --output <copy.pdf> [--force]
  officeimo pdf split <input.pdf> --output <new-folder> [--pages-per-document <count>] [--force]
  officeimo pdf decrypt <input.pdf> --password-env <owner-password-variable> --output <copy.pdf> [--force]
  officeimo pdf flatten <input.pdf> --output <copy.pdf> --acknowledge-raster-output
             [--pages <selection>] [--dpi 36..600] [--force]
  officeimo pdf ocr <input.pdf> --output <copy.pdf> --ocr-provider <id>
             --ocr-provider-assembly <provider.dll> [--ocr-language <language>]
             [--ocr-option <key=value>] [--ocr-min-confidence 0..1] [--pages <selection>] [--dpi 36..600] [--force]
  officeimo pdf optimize|sanitize <input.pdf> --output <copy.pdf> [--force]
  officeimo pdf convert-to-pdfa <input.pdf> --output <copy.pdf> [--profile 2b|2u|3b|3u|4|4f]
             [--attachments remove|refuse] [--font <family>[:bold|italic|bold-italic]=<font.ttf>]...
             [--font-substitution installed-fonts|none]
             [--output-intent <profile.icc> --output-condition-identifier <id>] [--password-env <name>] [--force]
  officeimo pdf providers [--ocr-provider-assembly <provider.dll>]
  officeimo pdf redact <command> [options]

Page selections preserve order and repeats and accept 1-3,last, odd, even, all, and ! exclusions.
Passwords are read only from --password-env, never from command-line values or result JSON.
All writes require a separate explicit output. Existing outputs are refused unless --force is supplied.
Flattening renders page appearances only: native text, forms, links, signatures, and attachments are omitted.
OCR adds searchable text to the source pages; recognition confidence is provider evidence, not a correctness guarantee.
PDF/A conversion writes a PDF/A copy, or refuses (exit 4) and writes nothing. Standard error lists every change or
violation; the JSON pdfA report keeps exact counts and is trimmed to the output budget (truncated=true).
--attachments defaults to remove; pass --attachments refuse to keep a source with embedded files unconverted.
Each supplied --font family needs a regular face.
An encrypted source needs its owner password. Supplied fonts are tried before installed fonts unless
--font-substitution none; a face is embedded only when its glyph widths match the page and its licence allows it.
Optional OCR provider assemblies and executable settings must be trusted by the invoking host.
Resource options: --maximum-pages 1..10000 (default100; split bounds parts),
  --maximum-input-bytes 1..268435456, --maximum-output-bytes 1..536870912,
  --maximum-pixels-per-page 1..100000000 (flatten/ocr; default25000000).
JSON results contain bounded artifact metadata and diagnostic codes, never document or OCR text.
""";

    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error,
        CancellationToken cancellationToken = default, OcrEngineCatalog? ocrCatalog = null) {
        try {
            PdfWorkflowArguments parsed = PdfWorkflowArguments.Parse(args);
            if (parsed.Help) { await output.WriteLineAsync(Usage).ConfigureAwait(false); return 0; }
            var catalog = ocrCatalog ?? new OcrEngineCatalog();
            PdfOcrProviderLoader.LoadExplicitAssemblies(catalog, parsed.ProviderAssemblies);
            var service = new OfficeImoAgentService(new AgentPathPolicy(), pdfOcrCatalog: catalog, pdfOcrProviderOptions: parsed.ProviderOptions,
                pdfPasswordEnvironmentVariables: parsed.Settings.PasswordEnvironmentVariable is { } variable ? [variable] : []);
            if (parsed.Operation == "providers") {
                await output.WriteLineAsync(AgentJson.Serialize(service.PdfOcrProviders().ToArray())).ConfigureAwait(false); return 0;
            }
            AgentPdfAReport? completeReport = null;
            AgentPdfWorkflowResult result = parsed.Operation switch {
                "split" => await service.SplitPdfAsync(parsed.Input!, parsed.Output!, parsed.PagesPerDocument, parsed.Settings, cancellationToken: cancellationToken).ConfigureAwait(false),
                "convert-to-pdfa" => await service.ConvertToPdfAAsync(parsed.Input!, parsed.Output!, parsed.PdfA, parsed.Settings,
                    maxOutputCharacters: 64000, completeReport: report => completeReport = report, cancellationToken: cancellationToken).ConfigureAwait(false),
                "ocr" => await service.SearchablePdfAsync(parsed.Input!, parsed.Output!, parsed.Provider!, parsed.Settings, parsed.Language, parsed.Confidence, cancellationToken: cancellationToken).ConfigureAwait(false),
                _ => await service.PdfAsync(parsed.Input!, parsed.Output!, parsed.Operation, parsed.Settings, cancellationToken: cancellationToken).ConfigureAwait(false)
            };
            await output.WriteLineAsync(AgentJson.Serialize(result)).ConfigureAwait(false);
            // Standard error lists the complete report; the JSON report is bounded with exact counts.
            if (completeReport is { } report) await WritePdfAReportAsync(report, error).ConfigureAwait(false);
            if (result.Succeeded) return 0;
            await error.WriteLineAsync("PDF " + parsed.Operation + " " + result.Status.ToLowerInvariant() + "; " + string.Join(", ", result.Diagnostics.Select(item => item.Code))).ConfigureAwait(false);
            return WorkflowCommand.MapStatus(Enum.Parse<OfficeWorkflowStatus>(result.Status), Enum.Parse<OfficeWorkflowFailureKind>(result.FailureKind));
        } catch (AgentUsageException exception) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false); return (int)OfficeImoToolExitCode.Usage;
        } catch (OperationCanceledException) {
            await error.WriteLineAsync("PDF workflow cancelled.").ConfigureAwait(false); return (int)OfficeImoToolExitCode.Cancelled;
        } catch (FileNotFoundException) {
            await error.WriteLineAsync("The input file or configured optional provider assembly was not found.").ConfigureAwait(false); return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (Exception exception) {
            // Optional providers own arbitrary messages; do not echo provider errors or secret configuration.
            await error.WriteLineAsync("PDF workflow failed: " + exception.GetType().Name).ConfigureAwait(false); return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }

    private static async Task WritePdfAReportAsync(AgentPdfAReport report, TextWriter error) {
        foreach (AgentPdfAEntry violation in report.Violations) await error.WriteLineAsync("violation " + Describe(violation)).ConfigureAwait(false);
        foreach (AgentPdfAEntry change in report.Changes) await error.WriteLineAsync("change " + Describe(change)).ConfigureAwait(false);
        int omitted = report.ChangeCount + report.ViolationCount - report.Changes.Count - report.Violations.Count;
        await error.WriteLineAsync(report.Profile + ": " + report.ChangeCount + " change(s), " + report.ViolationCount + " violation(s)" +
            (omitted > 0 ? "; " + omitted + " not listed" : string.Empty) + ".").ConfigureAwait(false);
    }

    private static string Describe(AgentPdfAEntry entry) => entry.Code +
        (entry.Location is null ? string.Empty : " at " + entry.Location) +
        (entry.ObjectNumber is int number && entry.Location != "object " + number ? " (object " + number + ")" : string.Empty) + ": " + entry.Message;
}
