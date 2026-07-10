using OfficeIMO.Reader;
using OfficeIMO.Reader.Ocr.Process;

namespace OfficeIMO.Reader.Ocr.Tesseract;

/// <summary>Optional cross-platform OCR engine backed by an installed Tesseract command-line executable.</summary>
public sealed partial class TesseractOcrEngine : IOfficeOcrEngine {
    private readonly OptionsSnapshot _options;
    private readonly OfficeOcrEngineCapabilities _capabilities;

    /// <summary>Creates a Tesseract engine from an immutable option snapshot.</summary>
    public TesseractOcrEngine(TesseractOcrEngineOptions? options = null) {
        _options = OptionsSnapshot.Create(options);
        _capabilities = new OfficeOcrEngineCapabilities {
            SupportedMediaTypes = new[] { "image/*" },
            SupportsLineSpans = true,
            SupportsWordSpans = true,
            SupportsCharacterSpans = false,
            SupportsConfidence = true,
            SupportsConcurrentRequests = true
        };
    }

    /// <inheritdoc />
    public string Id => "tesseract-cli";

    /// <inheritdoc />
    public OfficeOcrEngineCapabilities Capabilities => _capabilities.Clone();

    /// <inheritdoc />
    public async ValueTask<OfficeOcrEngineResult> RecognizeAsync(OfficeOcrEngineRequest request, CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (request.Payload == null || request.Payload.Length == 0) throw new ArgumentException("OCR request payload cannot be empty.", nameof(request));
        if (request.Payload.LongLength > _options.MaxInputBytes) throw new IOException("OCR request payload exceeds MaxInputBytes (" + _options.MaxInputBytes + ").");
        if (!string.IsNullOrWhiteSpace(request.Asset.MediaType) && !request.Asset.MediaType!.StartsWith("image/", StringComparison.OrdinalIgnoreCase)) throw new NotSupportedException("Tesseract provider accepts image assets only.");
        cancellationToken.ThrowIfCancellationRequested();
        string temporaryRoot = Path.GetFullPath(_options.TemporaryDirectory ?? Path.GetTempPath());
        string requestDirectory = OfficeOcrTemporaryStorage.CreateRequestDirectory(temporaryRoot, "officeimo-tesseract-");
        try {
            string inputPath = Path.Combine(requestDirectory, "input" + OfficeOcrProcessFileNames.GetSafeAssetExtension(request.Asset));
            string outputBase = Path.Combine(requestDirectory, "result");
            string outputPath = outputBase + ".tsv";
            OfficeOcrTemporaryStorage.WriteAllBytes(inputPath, request.Payload);
            string? language = string.IsNullOrWhiteSpace(request.Language) ? _options.Language : request.Language!.Trim();
            OfficeOcrProcessResult processResult = await OfficeOcrProcessRunner.RunAsync(new OfficeOcrProcessCommand {
                FileName = _options.ExecutablePath,
                Arguments = BuildRecognitionArguments(inputPath, outputBase, language),
                Timeout = _options.Timeout,
                MaxStandardOutputCharacters = _options.MaxProcessOutputCharacters,
                MaxStandardErrorCharacters = _options.MaxProcessOutputCharacters
            }, cancellationToken).ConfigureAwait(false);
            if (processResult.ExitCode != 0) {
                throw new InvalidOperationException("Tesseract exited with code " + processResult.ExitCode + ": " + processResult.StandardError);
            }
            if (!File.Exists(outputPath)) throw new FileNotFoundException("Tesseract did not create the expected TSV output.", outputPath);
            OfficeOcrTemporaryStorage.EnsurePrivateFile(outputPath);
            long outputLength = new FileInfo(outputPath).Length;
            if (outputLength > _options.MaxOutputBytes) throw new IOException("Tesseract TSV output exceeds MaxOutputBytes (" + _options.MaxOutputBytes + ").");
            OfficeOcrEngineResult result = TesseractTsvParser.Parse(File.ReadAllText(outputPath, Encoding.UTF8), language);
            if (!string.IsNullOrWhiteSpace(processResult.StandardError)) {
                result.Diagnostics = (result.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>()).Concat(new[] {
                    new OfficeDocumentDiagnostic {
                        Severity = OfficeDocumentDiagnosticSeverity.Warning,
                        Category = OfficeDocumentDiagnosticCategory.Ocr,
                        Code = "tesseract-stderr",
                        Message = processResult.StandardError,
                        Source = Id,
                        IsRecoverable = true,
                        Location = request.Candidate.Location,
                        Attributes = new Dictionary<string, string>(StringComparer.Ordinal) {
                            ["truncated"] = processResult.StandardErrorTruncated ? "true" : "false"
                        }
                    }
                }).ToArray();
            }
            if (processResult.StandardOutputTruncated) {
                result.Diagnostics = (result.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>()).Concat(new[] {
                    new OfficeDocumentDiagnostic {
                        Severity = OfficeDocumentDiagnosticSeverity.Warning,
                        Category = OfficeDocumentDiagnosticCategory.Limit,
                        Code = "tesseract-stdout-limit",
                        Message = "Tesseract standard output exceeded its configured retention limit.",
                        Source = Id,
                        IsRecoverable = true,
                        Location = request.Candidate.Location
                    }
                }).ToArray();
            }
            return result;
        } finally {
            if (!_options.KeepTemporaryFiles) TryDeleteDirectory(requestDirectory);
        }
    }

    internal IReadOnlyList<string> BuildRecognitionArguments(string inputPath, string outputBase, string? language) {
        var arguments = new List<string> { inputPath, outputBase };
        if (!string.IsNullOrWhiteSpace(language)) {
            arguments.Add("-l");
            arguments.Add(language!);
        }
        if (!string.IsNullOrWhiteSpace(_options.TessdataDirectory)) {
            arguments.Add("--tessdata-dir");
            arguments.Add(_options.TessdataDirectory!);
        }
        if (_options.EngineMode.HasValue) {
            arguments.Add("--oem");
            arguments.Add(_options.EngineMode.Value.ToString(CultureInfo.InvariantCulture));
        }
        if (_options.PageSegmentationMode.HasValue) {
            arguments.Add("--psm");
            arguments.Add(_options.PageSegmentationMode.Value.ToString(CultureInfo.InvariantCulture));
        }
        if (_options.Dpi.HasValue) {
            arguments.Add("--dpi");
            arguments.Add(_options.Dpi.Value.ToString(CultureInfo.InvariantCulture));
        }
        arguments.AddRange(_options.AdditionalArguments);
        arguments.Add("tsv");
        return arguments;
    }

    private static void TryDeleteDirectory(string path) {
        try {
            if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
        } catch (IOException) {
        } catch (UnauthorizedAccessException) {
        }
    }

    private sealed class OptionsSnapshot {
        internal string ExecutablePath { get; private set; } = string.Empty;
        internal string? Language { get; private set; }
        internal string? TessdataDirectory { get; private set; }
        internal int? EngineMode { get; private set; }
        internal int? PageSegmentationMode { get; private set; }
        internal int? Dpi { get; private set; }
        internal IReadOnlyList<string> AdditionalArguments { get; private set; } = Array.Empty<string>();
        internal string? TemporaryDirectory { get; private set; }
        internal TimeSpan Timeout { get; private set; }
        internal long MaxOutputBytes { get; private set; }
        internal long MaxInputBytes { get; private set; }
        internal int MaxProcessOutputCharacters { get; private set; }
        internal bool KeepTemporaryFiles { get; private set; }

        internal static OptionsSnapshot Create(TesseractOcrEngineOptions? options) {
            TesseractOcrEngineOptions source = options ?? new TesseractOcrEngineOptions();
            if (string.IsNullOrWhiteSpace(source.ExecutablePath)) throw new ArgumentException("Tesseract executable path cannot be empty.", nameof(options));
            if (source.EngineMode.HasValue && (source.EngineMode.Value < 0 || source.EngineMode.Value > 3)) throw new ArgumentOutOfRangeException(nameof(source.EngineMode));
            if (source.PageSegmentationMode.HasValue && (source.PageSegmentationMode.Value < 0 || source.PageSegmentationMode.Value > 13)) throw new ArgumentOutOfRangeException(nameof(source.PageSegmentationMode));
            if (source.Dpi.HasValue && source.Dpi.Value < 1) throw new ArgumentOutOfRangeException(nameof(source.Dpi));
            if (source.Timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(source.Timeout));
            if (source.MaxOutputBytes < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxOutputBytes));
            if (source.MaxInputBytes < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxInputBytes));
            if (source.MaxProcessOutputCharacters < 1) throw new ArgumentOutOfRangeException(nameof(source.MaxProcessOutputCharacters));
            return new OptionsSnapshot {
                ExecutablePath = source.ExecutablePath.Trim(),
                Language = string.IsNullOrWhiteSpace(source.Language) ? null : source.Language!.Trim(),
                TessdataDirectory = string.IsNullOrWhiteSpace(source.TessdataDirectory) ? null : source.TessdataDirectory,
                EngineMode = source.EngineMode,
                PageSegmentationMode = source.PageSegmentationMode,
                Dpi = source.Dpi,
                AdditionalArguments = (source.AdditionalArguments ?? Array.Empty<string>()).ToArray(),
                TemporaryDirectory = string.IsNullOrWhiteSpace(source.TemporaryDirectory) ? null : source.TemporaryDirectory,
                Timeout = source.Timeout,
                MaxOutputBytes = source.MaxOutputBytes,
                MaxInputBytes = source.MaxInputBytes,
                MaxProcessOutputCharacters = source.MaxProcessOutputCharacters,
                KeepTemporaryFiles = source.KeepTemporaryFiles
            };
        }
    }
}
