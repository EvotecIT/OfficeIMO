using System.Text;

namespace OfficeIMO.Reader.Tool;

internal static class ReaderToolOutput {
    private static readonly Encoding Utf8WithoutBom = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);

    internal static string FormatDocument(OfficeDocumentReadResult document, ReaderToolOutputFormat format) {
        if (format == ReaderToolOutputFormat.Json) {
            return OfficeDocumentReadResultJson.Serialize(document, indented: true);
        }

        if (!string.IsNullOrEmpty(document.Markdown)) {
            return document.Markdown!;
        }

        return string.Join(
            Environment.NewLine + Environment.NewLine,
            (document.Chunks ?? Array.Empty<ReaderChunk>())
                .Select(chunk => chunk.Markdown ?? chunk.Text)
                .Where(value => !string.IsNullOrWhiteSpace(value)));
    }

    internal static async Task WriteSingleAsync(
        string content,
        string? outputPath,
        TextWriter standardOutput,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(outputPath) || outputPath == "-") {
            await standardOutput.WriteAsync(content.AsMemory(), cancellationToken).ConfigureAwait(false);
            if (!content.EndsWith("\n", StringComparison.Ordinal)) {
                await standardOutput.WriteLineAsync().ConfigureAwait(false);
            }
            return;
        }

        await WriteFileAsync(outputPath!, content, cancellationToken).ConfigureAwait(false);
    }

    internal static async Task WriteFolderAsync(
        string sourceRoot,
        string outputRoot,
        string? assetsRoot,
        IReadOnlyList<string> paths,
        IReadOnlyList<OfficeDocumentReadResult> documents,
        ReaderToolOutputFormat format,
        CancellationToken cancellationToken) {
        if (File.Exists(outputRoot)) {
            throw new ReaderToolOutputException("Output path '" + outputRoot + "' is a file.");
        }

        try {
            Directory.CreateDirectory(outputRoot);
            if (!string.IsNullOrWhiteSpace(assetsRoot)) {
                Directory.CreateDirectory(assetsRoot!);
            }
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new ReaderToolOutputException("Could not create an output directory.", exception);
        }

        for (int index = 0; index < paths.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            string relativePath = Path.GetRelativePath(sourceRoot, paths[index]);
            string suffix = format == ReaderToolOutputFormat.Json ? ".reader.json" : ".md";
            string outputPath = Path.Combine(outputRoot, relativePath + suffix);
            ReaderToolPathSafety.EnsureOutsideInput(sourceRoot, outputPath);
            await WriteFileAsync(outputPath, FormatDocument(documents[index], format), cancellationToken)
                .ConfigureAwait(false);

            if (!string.IsNullOrWhiteSpace(assetsRoot) && documents[index].Assets.Count > 0) {
                string assetDirectory = Path.Combine(assetsRoot!, relativePath + ".assets");
                ReaderToolPathSafety.EnsureOutsideInput(sourceRoot, assetDirectory);
                WriteAssets(documents[index], assetDirectory, cancellationToken);
            }
        }
    }

    internal static void WriteAssets(
        OfficeDocumentReadResult document,
        string assetsPath,
        CancellationToken cancellationToken) {
        try {
            document.WriteAssetsToDirectory(
                assetsPath,
                new OfficeDocumentAssetMaterializationOptions {
                    CreateDirectory = true,
                    Overwrite = true,
                    ValidatePayloadHash = true
                },
                cancellationToken);
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new ReaderToolOutputException("Could not materialize document assets.", exception);
        }
    }

    private static async Task WriteFileAsync(string path, string content, CancellationToken cancellationToken) {
        string? temporaryPath = null;
        try {
            string fullPath = Path.GetFullPath(path);
            string? directory = Path.GetDirectoryName(fullPath);
            if (!string.IsNullOrEmpty(directory)) {
                Directory.CreateDirectory(directory);
            }
            string outputDirectory = string.IsNullOrEmpty(directory) ? Directory.GetCurrentDirectory() : directory!;
            temporaryPath = Path.Combine(
                outputDirectory,
                "." + Path.GetFileName(fullPath) + "." + Guid.NewGuid().ToString("N") + ".tmp");
            await File.WriteAllTextAsync(temporaryPath, content, Utf8WithoutBom, cancellationToken).ConfigureAwait(false);
            File.Move(temporaryPath, fullPath, overwrite: true);
            temporaryPath = null;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new ReaderToolOutputException("Could not write output file '" + path + "'.", exception);
        } finally {
            if (temporaryPath != null) {
                try {
                    File.Delete(temporaryPath);
                } catch (IOException) {
                } catch (UnauthorizedAccessException) {
                }
            }
        }
    }
}

internal sealed class ReaderToolOutputException : Exception {
    internal ReaderToolOutputException(string message) : base(message) { }
    internal ReaderToolOutputException(string message, Exception innerException) : base(message, innerException) { }
}
