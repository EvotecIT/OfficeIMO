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
  officeimo pdf providers [--ocr-provider-assembly <provider.dll>]
  officeimo pdf redact <command> [options]

Page selections preserve order and repeats and accept 1-3,last, odd, even, all, and ! exclusions.
Passwords are read only from --password-env, never from command-line values or result JSON.
All writes require a separate explicit output. Existing outputs are refused unless --force is supplied.
Flattening renders page appearances only: native text, forms, links, signatures, and attachments are omitted.
OCR adds searchable text to the source pages; recognition confidence is provider evidence, not a correctness guarantee.
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
            var service = new OfficeImoAgentService(new AgentPathPolicy(), pdfOcrCatalog: catalog, pdfOcrProviderOptions: parsed.ProviderOptions);
            if (parsed.Operation == "providers") {
                await output.WriteLineAsync(AgentJson.Serialize(service.PdfOcrProviders().ToArray())).ConfigureAwait(false); return 0;
            }
            AgentPdfWorkflowResult result = parsed.Operation switch {
                "split" => await service.SplitPdfAsync(parsed.Input!, parsed.Output!, parsed.PagesPerDocument, parsed.Settings, cancellationToken: cancellationToken).ConfigureAwait(false),
                "ocr" => await service.SearchablePdfAsync(parsed.Input!, parsed.Output!, parsed.Provider!, parsed.Settings, parsed.Language, parsed.Confidence, cancellationToken: cancellationToken).ConfigureAwait(false),
                _ => await service.PdfAsync(parsed.Input!, parsed.Output!, parsed.Operation, parsed.Settings, cancellationToken: cancellationToken).ConfigureAwait(false)
            };
            await output.WriteLineAsync(AgentJson.Serialize(result)).ConfigureAwait(false);
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
}
