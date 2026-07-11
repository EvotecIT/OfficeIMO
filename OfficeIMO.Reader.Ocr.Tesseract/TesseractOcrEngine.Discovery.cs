using OfficeIMO.Reader.Ocr.Process;

namespace OfficeIMO.Reader.Ocr.Tesseract;

public sealed partial class TesseractOcrEngine {
    /// <summary>Returns the first line from <c>tesseract --version</c>.</summary>
    public async Task<string> GetVersionAsync(CancellationToken cancellationToken = default) {
        OfficeOcrProcessResult result = await RunDiscoveryAsync(new[] { "--version" }, cancellationToken).ConfigureAwait(false);
        if (result.ExitCode != 0) throw new InvalidOperationException("Tesseract version discovery failed: " + result.StandardError);
        return result.StandardOutput.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries).FirstOrDefault()?.Trim() ?? string.Empty;
    }

    /// <summary>Returns sorted language identifiers from <c>tesseract --list-langs</c>.</summary>
    public async Task<IReadOnlyList<string>> GetLanguagesAsync(CancellationToken cancellationToken = default) {
        var arguments = new List<string> { "--list-langs" };
        if (!string.IsNullOrWhiteSpace(_options.TessdataDirectory)) {
            arguments.Add("--tessdata-dir");
            arguments.Add(_options.TessdataDirectory!);
        }
        OfficeOcrProcessResult result = await RunDiscoveryAsync(arguments, cancellationToken).ConfigureAwait(false);
        if (result.ExitCode != 0) throw new InvalidOperationException("Tesseract language discovery failed: " + result.StandardError);
        return result.StandardOutput
            .Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(static line => line.Trim())
            .Where(static line => line.Length > 0 && !line.StartsWith("List of available languages", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.Ordinal)
            .OrderBy(static line => line, StringComparer.Ordinal)
            .ToArray();
    }

    private Task<OfficeOcrProcessResult> RunDiscoveryAsync(IReadOnlyList<string> arguments, CancellationToken cancellationToken) {
        return OfficeOcrProcessRunner.RunAsync(new OfficeOcrProcessCommand {
            FileName = _options.ExecutablePath,
            Arguments = arguments,
            Timeout = _options.Timeout,
            MaxStandardOutputCharacters = _options.MaxProcessOutputCharacters,
            MaxStandardErrorCharacters = _options.MaxProcessOutputCharacters
        }, cancellationToken);
    }
}
