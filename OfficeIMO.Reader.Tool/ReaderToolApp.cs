using OfficeIMO.Reader.All;

namespace OfficeIMO.Reader.Tool;

internal static class ReaderToolApp {
    internal const string Usage = """
OfficeIMO.Reader.Tool

Usage:
  officeimo-reader read <path|-> [--name <source-name>] [--format markdown|json] [--output <file|->] [--assets <directory>]
                                [--max-input-bytes <bytes>]
  officeimo-reader folder <path> --output <directory> [--format markdown|json] [--assets <directory>] [--concurrency <1-64>]
                          [--max-files <count>] [--max-total-bytes <bytes>] [--max-input-bytes <bytes>] [--no-recursive]
  officeimo-reader capabilities [--format text|json]

The dependency-bounded tool does not configure OCR or hosted providers.
""";

    internal static async Task<int> RunAsync(
        string[] args,
        Stream standardInput,
        TextWriter standardOutput,
        TextWriter standardError,
        CancellationToken cancellationToken = default) {
        if (standardInput == null) throw new ArgumentNullException(nameof(standardInput));
        if (standardOutput == null) throw new ArgumentNullException(nameof(standardOutput));
        if (standardError == null) throw new ArgumentNullException(nameof(standardError));

        ReaderToolArguments parsed;
        try {
            parsed = ReaderToolArguments.Parse(args);
        } catch (ReaderToolUsageException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            await standardError.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)ReaderToolExitCode.Usage;
        }

        if (parsed.Command == ReaderToolCommand.Help) {
            await standardOutput.WriteLineAsync(Usage).ConfigureAwait(false);
            return (int)ReaderToolExitCode.Success;
        }

        try {
            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddAllOfficeIMOHandlers()
                .WithMaxConcurrentReads(parsed.Concurrency)
                .Build();

            return parsed.Command switch {
                ReaderToolCommand.Read => await RunReadAsync(
                    parsed, reader, standardInput, standardOutput, cancellationToken).ConfigureAwait(false),
                ReaderToolCommand.Folder => await RunFolderAsync(
                    parsed, reader, standardError, cancellationToken).ConfigureAwait(false),
                ReaderToolCommand.Capabilities => await RunCapabilitiesAsync(
                    parsed, reader, standardOutput).ConfigureAwait(false),
                _ => (int)ReaderToolExitCode.Usage
            };
        } catch (OperationCanceledException) {
            await standardError.WriteLineAsync("Operation cancelled.").ConfigureAwait(false);
            return (int)ReaderToolExitCode.Cancelled;
        } catch (ReaderToolOutputException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)ReaderToolExitCode.OutputFailed;
        } catch (FileNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)ReaderToolExitCode.InputNotFound;
        } catch (DirectoryNotFoundException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)ReaderToolExitCode.InputNotFound;
        } catch (NotSupportedException exception) {
            await standardError.WriteLineAsync(exception.Message).ConfigureAwait(false);
            return (int)ReaderToolExitCode.UnsupportedInput;
        } catch (Exception exception) {
            await standardError.WriteLineAsync(
                "Document read failed: " + exception.GetType().Name + ": " + exception.Message).ConfigureAwait(false);
            return (int)ReaderToolExitCode.ReadFailed;
        }
    }

    private static async Task<int> RunReadAsync(
        ReaderToolArguments options,
        OfficeDocumentReader reader,
        Stream standardInput,
        TextWriter standardOutput,
        CancellationToken cancellationToken) {
        var readerOptions = new ReaderOptions { MaxInputBytes = options.MaxInputBytes };
        OfficeDocumentReadResult document;
        if (options.InputPath == "-") {
            document = await reader.ReadDocumentAsync(
                standardInput,
                options.SourceName,
                readerOptions,
                cancellationToken: cancellationToken).ConfigureAwait(false);
        } else {
            string inputPath = Path.GetFullPath(options.InputPath!);
            if (!File.Exists(inputPath)) {
                throw new FileNotFoundException("Input file '" + inputPath + "' does not exist.", inputPath);
            }
            if (!string.IsNullOrWhiteSpace(options.OutputPath) && options.OutputPath != "-") {
                ReaderToolPathSafety.EnsureDistinctFile(inputPath, options.OutputPath!);
            }
            document = await reader.ReadDocumentAsync(inputPath, readerOptions, cancellationToken)
                .ConfigureAwait(false);
        }

        await ReaderToolOutput.WriteSingleAsync(
            ReaderToolOutput.FormatDocument(document, options.Format),
            options.OutputPath,
            standardOutput,
            cancellationToken).ConfigureAwait(false);
        if (!string.IsNullOrWhiteSpace(options.AssetsPath)) {
            ReaderToolOutput.WriteAssets(document, options.AssetsPath!, cancellationToken);
        }
        return (int)ReaderToolExitCode.Success;
    }

    private static async Task<int> RunFolderAsync(
        ReaderToolArguments options,
        OfficeDocumentReader reader,
        TextWriter standardError,
        CancellationToken cancellationToken) {
        string inputPath = Path.GetFullPath(options.InputPath!);
        if (!Directory.Exists(inputPath)) {
            throw new DirectoryNotFoundException("Input folder '" + inputPath + "' does not exist.");
        }

        string outputPath = Path.GetFullPath(options.OutputPath!);
        string? assetsPath = string.IsNullOrWhiteSpace(options.AssetsPath)
            ? null
            : Path.GetFullPath(options.AssetsPath!);
        ReaderToolPathSafety.EnsureOutsideInput(inputPath, outputPath, assetsPath);

        IReadOnlyList<string> paths = ReaderToolFileDiscovery.FindSupportedFiles(
            inputPath,
            reader,
            options.Recurse,
            options.MaxFiles,
            options.MaxTotalBytes,
            cancellationToken);
        IReadOnlyList<OfficeDocumentReadResult> documents = await reader.ReadDocumentsAsync(
            paths,
            options: new ReaderOptions { MaxInputBytes = options.MaxInputBytes },
            batchOptions: new ReaderBatchOptions {
                MaxDegreeOfParallelism = options.Concurrency,
                MaxDocuments = options.MaxFiles
            },
            cancellationToken: cancellationToken).ConfigureAwait(false);

        await ReaderToolOutput.WriteFolderAsync(
            inputPath,
            outputPath,
            assetsPath,
            paths,
            documents,
            options.Format,
            cancellationToken).ConfigureAwait(false);
        await standardError.WriteLineAsync("Converted " + paths.Count + " document(s).").ConfigureAwait(false);
        return (int)ReaderToolExitCode.Success;
    }

    private static async Task<int> RunCapabilitiesAsync(
        ReaderToolArguments options,
        OfficeDocumentReader reader,
        TextWriter standardOutput) {
        if (options.Format == ReaderToolOutputFormat.Json) {
            await standardOutput.WriteLineAsync(reader.GetCapabilityManifestJson(indented: true)).ConfigureAwait(false);
            return (int)ReaderToolExitCode.Success;
        }

        foreach (ReaderHandlerCapability capability in reader.GetCapabilities()) {
            await standardOutput.WriteLineAsync(
                capability.Id + "\t" + capability.Origin + "\t" + string.Join(",", capability.Extensions)).ConfigureAwait(false);
        }
        return (int)ReaderToolExitCode.Success;
    }
}
