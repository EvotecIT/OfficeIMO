using System.Globalization;
using OfficeIMO.Pdf;
using OfficeIMO.Workflows;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal AgentPdfWorkflowResult ExtractPdfAssets(IReadOnlyList<string> paths, string outputDirectory,
        bool compositeSoftMasks, PdfWorkflowSettings settings, int maxOutputCharacters = 4000,
        CancellationToken cancellationToken = default) {
        settings.Validate();
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        string[] inputs = paths.Select(ResolvePdfInput).ToArray();
        string destination = PreparePdfOutput(outputDirectory, inputs, true, false, "extract-assets", maxOutputCharacters);
        var loadOptions = new PdfLoadOptions {
            Password = PdfPassword(settings),
            Limits = new PdfReadLimits { MaxInputBytes = settings.MaximumInputBytes, MaxPages = settings.MaximumPages }
        };
        var extractionOptions = new PdfAssetExtractionOptions {
            CompositeSoftMasks = compositeSoftMasks,
            MaximumPages = settings.MaximumPages,
            MaximumOutputBytes = settings.MaximumOutputBytes
        };

        IReadOnlyList<string> files = PdfAssetOutput.WriteDirectory(destination, create => {
            var result = new List<string>();
            for (int index = 0; index < inputs.Length; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                using Stream input = File.OpenRead(_pathPolicy.ResolveInput(inputs[index]));
                using PdfDocument document = PdfDocument.Load(input, loadOptions, cancellationToken);
                string prefix = inputs.Length == 1 ? string.Empty : (index + 1).ToString("D3", CultureInfo.InvariantCulture) + "/";
                IReadOnlyList<string> bundle = document.ExtractAssetsTo(name => create(prefix + name), extractionOptions, cancellationToken);
                result.AddRange(bundle.Select(name => prefix + name));
            }
            return result;
        }, ValidateOutput, cancellationToken);
        AgentPdfArtifact[] artifacts = files
            .OrderByDescending(static name => Path.GetFileName(name) is "report.json" or "report.html")
            .Select(name => {
                string file = Path.Combine(destination, name.Replace('/', Path.DirectorySeparatorChar));
                return new AgentPdfArtifact { Path = file, SizeBytes = new FileInfo(file).Length };
            }).ToArray();
        return PdfResult("extract-assets", OfficeWorkflowStatus.Completed, OfficeWorkflowFailureKind.None, destination,
            artifacts.Sum(static artifact => artifact.SizeBytes), artifacts, [], maxOutputCharacters);

        void ValidateOutput(string pathToCheck) {
            if (!string.Equals(pathToCheck, _pathPolicy.ResolveOutput(pathToCheck), StringComparison.Ordinal)) {
                throw new UnauthorizedAccessException("The asset output path changed during publication.");
            }
        }
    }
}
