using OfficeIMO.Tool.Agent;

namespace OfficeIMO.Tool.Commands.Pdf;

internal static class PdfRedlineCommand {
    internal const string Usage = """
OfficeIMO.Tool - PDF text redline

Usage:
  officeimo pdf redline <old.pdf> <new.pdf> --output <annotated.pdf>
             [--changes <changes.json>] [--report <report.html>]
             [--algorithm myers|patience|histogram] [--ignore-case]
             [--password-env <name>] [--new-password-env <name>] [--force]

Compares the text of two PDFs word by word in reading order and writes a copy of the new PDF with review
annotations: highlights on inserted, replaced and moved text, and a caret where text was deleted. Notes on
replacements, deletions and moves carry the old text. --changes writes every change as JSON with page numbers and
page rectangles; --report writes a side-by-side HTML comparison. PDFs without a text layer are refused (exit 2);
add one with 'officeimo pdf ocr' first. --password-env opens both PDFs; --new-password-env overrides it for the new PDF.
Existing outputs are refused unless --force is supplied. The JSON result on standard output has counts and
artifact paths, never document text.
""";

    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken cancellationToken = default) {
        try {
            var positional = new List<string>();
            string? outputPath = null, changes = null, report = null, password = null, newPassword = null, algorithm = "myers";
            bool ignoreCase = false, force = false;
            for (int index = 0; index < args.Length; index++) {
                string argument = args[index];
                string Value() => index + 1 < args.Length ? args[++index] : throw new AgentUsageException(argument + " requires a value.");
                switch (argument.ToLowerInvariant()) {
                    case "--help" or "-h" or "-?": await output.WriteLineAsync(Usage).ConfigureAwait(false); return (int)OfficeImoToolExitCode.Success;
                    case "--output" or "-o": outputPath = Value(); break;
                    case "--changes": changes = Value(); break;
                    case "--report": report = Value(); break;
                    case "--algorithm": algorithm = Value(); break;
                    case "--password-env": password = Value(); break;
                    case "--new-password-env": newPassword = Value(); break;
                    case "--ignore-case": ignoreCase = true; break;
                    case "--force": force = true; break;
                    default:
                        if (argument.StartsWith("-", StringComparison.Ordinal)) throw new AgentUsageException("Unknown option " + argument + ".");
                        positional.Add(argument);
                        break;
                }
            }

            if (positional.Count != 2) throw new AgentUsageException("Supply the old PDF and the new PDF.");
            if (outputPath is null) throw new AgentUsageException("--output is required.");
            string[] admitted = new[] { password, newPassword }.Where(static name => name is not null).Select(static name => name!).ToArray();
            var service = new OfficeImoAgentService(new AgentPathPolicy(), pdfPasswordEnvironmentVariables: admitted);
            var settings = new PdfWorkflowSettings { PasswordEnvironmentVariable = password, Overwrite = force };
            AgentPdfWorkflowResult result = await service.RedlinePdfAsync(positional[0], positional[1], outputPath, changes, report, settings,
                newPassword, ignoreCase, algorithm, maxOutputCharacters: 64000, cancellationToken).ConfigureAwait(false);
            await output.WriteLineAsync(AgentJson.Serialize(result)).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Success;
        } catch (AgentUsageException exception) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Usage;
        } catch (OperationCanceledException) {
            await error.WriteLineAsync("PDF redline cancelled.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.Cancelled;
        } catch (FileNotFoundException) {
            await error.WriteLineAsync("An input PDF was not found.").ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (Exception exception) {
            await error.WriteLineAsync("PDF redline failed: " + exception.GetType().Name).ConfigureAwait(false);
            return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }
}
