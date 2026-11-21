using OfficeIMO.AsciiDoc;
using OfficeIMO.Latex;
using OfficeIMO.OneNote;
using OfficeIMO.OpenDocument;
using OfficeIMO.Rtf;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    private sealed partial class PairRoute {
        public async Task<bool> TryValidateArtifactAsync(Stream output, OfficeConversionOptions options, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            output.Position = 0;
            OfficeWorkflowRouteSettings settings = CaptureSettings(options);
            string extension = Path.GetExtension(options.TargetName!).ToLowerInvariant();
            switch (extension) {
                case ".adoc":
                case ".asciidoc": {
                    var parsed = AsciiDocDocument.LoadResult(output, cancellationToken: token, options: null, encoding: null);
                    if (parsed.Document.Diagnostics.Any(d => d.Severity == AsciiDocDiagnosticSeverity.Error))
                        throw new InvalidDataException("Converted AsciiDoc contains native parse errors.");
                    break;
                }
                case ".tex":
                case ".latex": {
                    var parsed = LatexDocument.LoadResult(output, options: null, encoding: null, cancellationToken: token);
                    if (parsed.Document.Diagnostics.Any(d => d.Severity == LatexDiagnosticSeverity.Error))
                        throw new InvalidDataException("Converted LaTeX contains native parse errors.");
                    break;
                }
                case ".rtf": {
                    var parsed = await RtfDocument.LoadResultAsync(output, cancellationToken: token).ConfigureAwait(false);
                    using RtfDocument document = parsed.Document;
                    if (parsed.Diagnostics.Any(d => d.Severity == OfficeIMO.Rtf.Diagnostics.RtfDiagnosticSeverity.Error))
                        throw new InvalidDataException("Converted RTF contains native parse errors.");
                    break;
                }
                case ".odt":
                case ".ods":
                case ".odp": {
                    var load = new OdfLoadOptions { MaxPackageBytes = options.MaximumOutputBytes,
                        MaxXmlCharacters = options.MaximumXmlCharactersInPart, Password = settings.OpenDocumentSave?.Encryption?.Password };
                    using OdfDocument document = extension switch {
                        ".odt" => await OdtDocument.LoadAsync(output, load, token).ConfigureAwait(false),
                        ".ods" => await OdsDocument.LoadAsync(output, load, token).ConfigureAwait(false),
                        _ => await OdpPresentation.LoadAsync(output, load, token).ConfigureAwait(false)
                    };
                    break;
                }
                case ".csv": {
                    using OfficeIMO.CSV.CsvDocument csv = await OfficeIMO.CSV.CsvDocument.LoadAsync(output, null, token).ConfigureAwait(false);
                    break;
                }
                case ".one": {
                    using OneNoteSection section = OneNoteSection.Load(output,
                        new OneNoteLoadOptions { MaxInputBytes = options.MaximumOutputBytes });
                    break;
                }
                default: return false;
            }
            token.ThrowIfCancellationRequested();
            output.Position = output.Length;
            return true;
        }
    }
}
