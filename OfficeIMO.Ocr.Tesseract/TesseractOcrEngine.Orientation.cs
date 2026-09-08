using OfficeIMO.Ocr;
using OfficeIMO.Ocr.Process;

namespace OfficeIMO.Ocr.Tesseract;

public sealed partial class TesseractOcrEngine {
    private async Task<OcrResult> DetectOrientationAsync(OcrRequest request, CancellationToken cancellationToken) {
        string temporaryRoot = Path.GetFullPath(_options.TemporaryDirectory ?? Path.GetTempPath());
        string directory = OcrTemporaryStorage.CreateRequestDirectory(temporaryRoot, "officeimo-tesseract-osd-");
        try {
            string input = Path.Combine(directory, "input" + OcrProcessFileNames.GetSafeExtension(request.FileName, request.MediaType));
            OcrTemporaryStorage.WriteAllBytes(input, request.Payload);
            var arguments = new List<string> { input, "stdout", "-l", "osd", "--psm", "0" };
            if (!string.IsNullOrWhiteSpace(_options.TessdataDirectory)) {
                arguments.Add("--tessdata-dir"); arguments.Add(_options.TessdataDirectory!);
            }
            if (_options.Dpi.HasValue) {
                arguments.Add("--dpi"); arguments.Add(_options.Dpi.Value.ToString(CultureInfo.InvariantCulture));
            }
            OcrProcessResult process = await OcrProcessRunner.RunAsync(new OcrProcessCommand {
                FileName = _options.ExecutablePath, Arguments = arguments, Timeout = _options.Timeout,
                MaxStandardOutputCharacters = _options.MaxProcessOutputCharacters,
                MaxStandardErrorCharacters = _options.MaxProcessOutputCharacters
            }, cancellationToken).ConfigureAwait(false);
            OcrOrientationResult? orientation = process.ExitCode == 0 && !process.StandardOutputTruncated
                ? ParseOrientation(process.StandardOutput) : null;
            var diagnostics = new List<OcrDiagnostic>();
            if (orientation == null) diagnostics.Add(new OcrDiagnostic {
                Code = "tesseract-orientation-unavailable", Severity = OcrDiagnosticSeverity.Warning,
                Message = "Orientation could not be established. Retain the original orientation; verify osd trained data and sufficient text.",
                Source = Id, IsRecoverable = true
            });
            if (!string.IsNullOrWhiteSpace(process.StandardError)) diagnostics.Add(new OcrDiagnostic {
                Code = "tesseract-orientation-stderr", Severity = OcrDiagnosticSeverity.Warning,
                Message = process.StandardError, Source = Id, IsRecoverable = true
            });
            return new OcrResult { Provider = Id, Orientation = orientation, Diagnostics = diagnostics.AsReadOnly() };
        } finally {
            if (!_options.KeepTemporaryFiles) TryDeleteDirectory(directory);
        }
    }

    internal static OcrOrientationResult? ParseOrientation(string text) {
        int? rotation = null;
        double? confidence = null;
        string? script = null;
        foreach (string line in text.Replace("\r", "").Split('\n')) {
            int separator = line.IndexOf(':');
            if (separator < 0) continue;
            string name = line.Substring(0, separator).Trim(), value = line.Substring(separator + 1).Trim();
            if (name == "Rotate") {
                if (rotation.HasValue || !int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedRotation)) return null;
                rotation = parsedRotation;
            } else if (name == "Orientation confidence") {
                if (confidence.HasValue || !double.TryParse(value, NumberStyles.Float, CultureInfo.InvariantCulture, out double parsedConfidence)) return null;
                confidence = parsedConfidence;
            }
            else if (name == "Script" && value.Length <= 256) script = value;
        }
        if (rotation is not (0 or 90 or 180 or 270) || !confidence.HasValue ||
            double.IsNaN(confidence.Value) || double.IsInfinity(confidence.Value) || confidence.Value < 0D) return null;
        return new OcrOrientationResult {
            ClockwiseRotationDegrees = rotation.Value,
            // Tesseract documents 15 as reasonably confident; this is a bounded provider-specific scale, not a probability.
            Confidence = Math.Min(1D, confidence.Value / 15D), Script = script
        };
    }
}
