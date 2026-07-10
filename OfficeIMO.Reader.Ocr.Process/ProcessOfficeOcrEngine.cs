using OfficeIMO.Reader;

namespace OfficeIMO.Reader.Ocr.Process;

/// <summary>Runs an external executable that implements the OfficeIMO OCR JSON file protocol.</summary>
public sealed class ProcessOfficeOcrEngine : IOfficeOcrEngine {
    private readonly OptionsSnapshot _options;

    /// <summary>Creates a process-backed OCR engine from an immutable option snapshot.</summary>
    public ProcessOfficeOcrEngine(ProcessOfficeOcrEngineOptions options) {
        _options = OptionsSnapshot.Create(options);
    }

    /// <inheritdoc />
    public string Id => _options.Id;

    /// <inheritdoc />
    public OfficeOcrEngineCapabilities Capabilities => _options.Capabilities.Clone();

    /// <inheritdoc />
    public async ValueTask<OfficeOcrEngineResult> RecognizeAsync(OfficeOcrEngineRequest request, CancellationToken cancellationToken = default) {
        if (request == null) throw new ArgumentNullException(nameof(request));
        if (request.Payload == null || request.Payload.Length == 0) throw new ArgumentException("OCR request payload cannot be empty.", nameof(request));
        if (request.Payload.LongLength > _options.MaxInputBytes) throw new IOException("OCR request payload exceeds MaxInputBytes (" + _options.MaxInputBytes + ").");
        cancellationToken.ThrowIfCancellationRequested();
        string temporaryRoot = Path.GetFullPath(_options.TemporaryDirectory ?? Path.GetTempPath());
        string requestDirectory = Path.Combine(temporaryRoot, "officeimo-ocr-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(requestDirectory);
        try {
            string inputPath = Path.Combine(requestDirectory, "input" + OfficeOcrProcessFileNames.GetSafeAssetExtension(request.Asset));
            string requestPath = Path.Combine(requestDirectory, "request.json");
            string outputPath = Path.Combine(requestDirectory, "result.json");
            File.WriteAllBytes(inputPath, request.Payload);
            var processRequest = new ProcessOfficeOcrRequest {
                CandidateId = request.Candidate.Id,
                CandidateKind = request.Candidate.Kind,
                AssetId = request.Asset.Id,
                MediaType = request.Asset.MediaType,
                SourcePath = request.Source.Path,
                InputPath = inputPath,
                OutputPath = outputPath,
                Language = request.Language,
                Location = request.Candidate.Location,
                Region = request.Candidate.Region,
                ProviderOptions = request.ProviderOptions
            };
            File.WriteAllText(requestPath, ProcessOfficeOcrProtocol.SerializeRequest(processRequest, indented: true), new UTF8Encoding(false));
            IReadOnlyList<string> arguments = _options.Arguments.Select(argument => ExpandArgument(argument, processRequest, requestPath)).ToArray();
            OfficeOcrProcessResult processResult = await OfficeOcrProcessRunner.RunAsync(new OfficeOcrProcessCommand {
                FileName = _options.FileName,
                Arguments = arguments,
                WorkingDirectory = _options.WorkingDirectory,
                EnvironmentVariables = _options.EnvironmentVariables,
                Timeout = _options.Timeout,
                MaxStandardOutputCharacters = _options.MaxProcessOutputCharacters,
                MaxStandardErrorCharacters = _options.MaxProcessOutputCharacters
            }, cancellationToken).ConfigureAwait(false);
            if (processResult.ExitCode != 0) {
                throw new InvalidOperationException("OCR process exited with code " + processResult.ExitCode + ": " + processResult.StandardError);
            }
            if (!File.Exists(outputPath)) throw new FileNotFoundException("OCR process did not create its response file.", outputPath);
            long outputLength = new FileInfo(outputPath).Length;
            if (outputLength > _options.MaxOutputBytes) throw new IOException("OCR process response exceeds MaxOutputBytes (" + _options.MaxOutputBytes + ").");
            OfficeOcrEngineResult result = ProcessOfficeOcrProtocol.DeserializeResult(File.ReadAllText(outputPath, Encoding.UTF8));
            if (string.IsNullOrWhiteSpace(result.Provider)) result.Provider = Id;
            if (!string.IsNullOrWhiteSpace(processResult.StandardError)) {
                result.Diagnostics = (result.Diagnostics ?? Array.Empty<OfficeDocumentDiagnostic>()).Concat(new[] {
                    new OfficeDocumentDiagnostic {
                        Severity = OfficeDocumentDiagnosticSeverity.Warning,
                        Category = OfficeDocumentDiagnosticCategory.Ocr,
                        Code = "ocr-process-stderr",
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
                        Code = "ocr-process-stdout-limit",
                        Message = "OCR process standard output exceeded its configured retention limit.",
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

    private static string ExpandArgument(string value, ProcessOfficeOcrRequest request, string requestPath) {
        return (value ?? string.Empty)
            .Replace("{request}", requestPath)
            .Replace("{input}", request.InputPath)
            .Replace("{output}", request.OutputPath)
            .Replace("{language}", request.Language ?? string.Empty)
            .Replace("{candidateId}", request.CandidateId)
            .Replace("{assetId}", request.AssetId);
    }

    private static void TryDeleteDirectory(string path) {
        try {
            if (Directory.Exists(path)) Directory.Delete(path, recursive: true);
        } catch (IOException) {
        } catch (UnauthorizedAccessException) {
        }
    }

    private sealed class OptionsSnapshot {
        internal string FileName { get; private set; } = string.Empty;
        internal IReadOnlyList<string> Arguments { get; private set; } = Array.Empty<string>();
        internal string Id { get; private set; } = string.Empty;
        internal string? WorkingDirectory { get; private set; }
        internal IReadOnlyDictionary<string, string> EnvironmentVariables { get; private set; } = new Dictionary<string, string>();
        internal string? TemporaryDirectory { get; private set; }
        internal TimeSpan Timeout { get; private set; }
        internal long MaxOutputBytes { get; private set; }
        internal long MaxInputBytes { get; private set; }
        internal int MaxProcessOutputCharacters { get; private set; }
        internal bool KeepTemporaryFiles { get; private set; }
        internal OfficeOcrEngineCapabilities Capabilities { get; private set; } = new OfficeOcrEngineCapabilities();

        internal static OptionsSnapshot Create(ProcessOfficeOcrEngineOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
            if (string.IsNullOrWhiteSpace(options.FileName)) throw new ArgumentException("OCR process filename cannot be empty.", nameof(options));
            if (string.IsNullOrWhiteSpace(options.Id)) throw new ArgumentException("OCR process id cannot be empty.", nameof(options));
            if (options.Timeout <= TimeSpan.Zero) throw new ArgumentOutOfRangeException(nameof(options.Timeout));
            if (options.MaxOutputBytes < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxOutputBytes));
            if (options.MaxInputBytes < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxInputBytes));
            if (options.MaxProcessOutputCharacters < 1) throw new ArgumentOutOfRangeException(nameof(options.MaxProcessOutputCharacters));
            OfficeOcrEngineCapabilities capabilities = options.Capabilities ?? new OfficeOcrEngineCapabilities();
            return new OptionsSnapshot {
                FileName = options.FileName.Trim(),
                Arguments = (options.Arguments ?? Array.Empty<string>()).ToArray(),
                Id = options.Id.Trim(),
                WorkingDirectory = string.IsNullOrWhiteSpace(options.WorkingDirectory) ? null : options.WorkingDirectory,
                EnvironmentVariables = options.EnvironmentVariables == null
                    ? new Dictionary<string, string>(StringComparer.Ordinal)
                    : options.EnvironmentVariables.ToDictionary(static pair => pair.Key, static pair => pair.Value, StringComparer.Ordinal),
                TemporaryDirectory = string.IsNullOrWhiteSpace(options.TemporaryDirectory) ? null : options.TemporaryDirectory,
                Timeout = options.Timeout,
                MaxOutputBytes = options.MaxOutputBytes,
                MaxInputBytes = options.MaxInputBytes,
                MaxProcessOutputCharacters = options.MaxProcessOutputCharacters,
                KeepTemporaryFiles = options.KeepTemporaryFiles,
                Capabilities = new OfficeOcrEngineCapabilities {
                    SupportedMediaTypes = (capabilities.SupportedMediaTypes ?? Array.Empty<string>()).ToArray(),
                    SupportedLanguages = (capabilities.SupportedLanguages ?? Array.Empty<string>()).ToArray(),
                    SupportsLineSpans = capabilities.SupportsLineSpans,
                    SupportsWordSpans = capabilities.SupportsWordSpans,
                    SupportsCharacterSpans = capabilities.SupportsCharacterSpans,
                    SupportsConfidence = capabilities.SupportsConfidence,
                    SupportsConcurrentRequests = capabilities.SupportsConcurrentRequests
                }
            };
        }
    }
}
