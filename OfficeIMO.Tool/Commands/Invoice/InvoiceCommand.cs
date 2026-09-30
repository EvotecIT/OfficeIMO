using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Commands.Invoice;

internal static class InvoiceCommand {
    internal const string Usage = """
OfficeIMO.Tool - invoice workflows

Usage:
  officeimo invoice inspect|validate <input.xml> [options]
  officeimo invoice convert|render|hybrid <input.xml> --output <path> <target options> [options]
  officeimo invoice batch inspect|validate <input.xml>... [options]
  officeimo invoice batch convert|render|hybrid <input.xml>... --output-directory <path> <target options> [options]
  officeimo invoice edit <input.xml> --output <path> <field replacements> [options]
  officeimo invoice batch edit <input.xml>... --output-directory <path> <field replacements> [options]

Target options (all three required for convert/render/hybrid):
  --release En16931_1_3_16|FacturX_1_09_2_Zugferd_2_5_2|XRechnung_3_0_2_2026_08_31|PeppolBis_3_0_21
  --syntax Cii|Ubl
  --profile Minimum|BasicWithoutLines|Basic|En16931|Extended|XRechnung|PeppolBis
  --allow-profile-loss                  Allow documented lower-profile reductions, reported as warnings

Source-field replacements (one or more, only for edit):
  --number <value> --issue-date <yyyy-MM-dd> --due-date <yyyy-MM-dd>
  --buyer-reference <value> --payment-reference <value>
  Edits require existing unique plaintext fields and retain other XML content.
  Signed XML cannot be edited. Completion alone does not establish model validity.

Standards validation:
  --standards-release <release> --rule-bundle <pinned-xrechnung.zip>
  --peppol-rules <pinned-peppol.sch> --facturx-rules <pinned-facturx.zip>
  --saxon-jar <saxon-he.jar> [--java <executable>]

Rendering and batches:
  --language <en-US,pl-PL,...> --font <ttf-or-otf>
  --columns <Item,Quantity,Unit,NetPrice,Vat,NetAmount>   Two to eight distinct columns including Item and NetAmount
  --unit-display Code|Description|CodeAndDescription --payment-display Code|Description|CodeAndDescription
  --modern --compact-details --page-identity
  --max-items <1-10000> --max-input-bytes <bytes> --max-output-bytes <bytes> --stop-on-failure

JSON carries model, mapping, target and exact-byte standards evidence separately.
Standards stages remain NotRun unless explicitly requested and configured.
Rendering requires CII. 'hybrid' embeds the captured XML; 'render' produces a separate presentation.
Outputs are created atomically and never overwrite existing files. Batch publication is per item.
""";

    internal static async Task<int> RunAsync(string[] args, TextWriter output, TextWriter error, CancellationToken token) {
        try {
            InvoiceArguments arguments = InvoiceArguments.Parse(args);
            if (arguments.Help) { await output.WriteLineAsync(Usage).ConfigureAwait(false); return 0; }
            var results = await OfficeInvoiceFileWorkflow.RunBatchAsync(arguments.Requests(), arguments.Limits, arguments.Validator(), token).ConfigureAwait(false);
            await InvoiceOutput.WriteAsync(output, results, arguments.Batch).ConfigureAwait(false);
            return results.Any(r => r.PublicationError != null) ? (int)OfficeImoToolExitCode.OutputFailed
                : results.Any(r => !r.Succeeded) ? (int)OfficeImoToolExitCode.ValidationFailed : (int)OfficeImoToolExitCode.Success;
        } catch (OperationCanceledException) { await error.WriteLineAsync("Invoice workflow cancelled.").ConfigureAwait(false); return (int)OfficeImoToolExitCode.Cancelled;
        } catch (OfficeInvoiceOutputException exception) { await error.WriteLineAsync(exception.Message).ConfigureAwait(false); return (int)OfficeImoToolExitCode.OutputFailed;
        } catch (FileNotFoundException exception) { await error.WriteLineAsync(exception.Message).ConfigureAwait(false); return (int)OfficeImoToolExitCode.InputNotFound;
        } catch (Exception exception) when (exception is ArgumentException or System.Xml.XmlException) { await error.WriteLineAsync(exception.Message).ConfigureAwait(false); return (int)OfficeImoToolExitCode.Usage;
        } catch (Exception exception) when (exception is IOException or InvalidDataException or UnauthorizedAccessException or NotSupportedException or InvalidOperationException) {
            await error.WriteLineAsync(exception.Message).ConfigureAwait(false); return (int)OfficeImoToolExitCode.OperationFailed;
        }
    }
}
