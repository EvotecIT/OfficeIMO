using OfficeIMO.Reader;
using System.Text;

namespace OfficeIMO.Tool.Commands.Reader;

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
        bool overwrite,
        CancellationToken cancellationToken) {
        if (string.IsNullOrWhiteSpace(outputPath) || outputPath == "-") {
            await standardOutput.WriteAsync(content.AsMemory(), cancellationToken).ConfigureAwait(false);
            if (!content.EndsWith("\n", StringComparison.Ordinal)) {
                await standardOutput.WriteLineAsync().ConfigureAwait(false);
            }
            return;
        }

        await WriteFileAsync(outputPath!, content, overwrite, cancellationToken).ConfigureAwait(false);
    }

    internal static async Task WriteSingleDocumentAsync(
        OfficeDocumentReadResult document,
        ReaderToolOutputFormat format,
        string? outputPath,
        TextWriter standardOutput,
        string? assetsPath,
        string? sourcePath,
        bool overwrite,
        CancellationToken cancellationToken) {
        if (!string.IsNullOrWhiteSpace(assetsPath)) {
            PrepareAssetsOutput(document, assetsPath!, overwrite, outputPath, sourcePath);
        }

        await WriteSingleAsync(
            FormatDocument(document, format),
            outputPath,
            standardOutput,
            overwrite,
            cancellationToken).ConfigureAwait(false);

        if (!string.IsNullOrWhiteSpace(assetsPath)) {
            WriteAssets(document, assetsPath!, overwrite, cancellationToken, outputPath, sourcePath);
        }
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

            string? assetDirectory = null;
            if (!string.IsNullOrWhiteSpace(assetsRoot) && documents[index].Assets.Count > 0) {
                assetDirectory = Path.Combine(assetsRoot!, relativePath + ".assets");
                ReaderToolPathSafety.EnsureOutsideInput(sourceRoot, assetDirectory);
                PrepareAssetsOutput(
                    documents[index],
                    assetDirectory,
                    overwrite: true,
                    outputPath,
                    paths[index]);
            }

            await WriteFileAsync(
                    outputPath,
                    FormatDocument(documents[index], format),
                    overwrite: true,
                    cancellationToken)
                .ConfigureAwait(false);

            if (assetDirectory != null) {
                WriteAssets(
                    documents[index],
                    assetDirectory,
                    overwrite: true,
                    cancellationToken,
                    outputPath,
                    paths[index]);
            }
        }
    }

    internal static void WriteAssets(
        OfficeDocumentReadResult document,
        string assetsPath,
        bool overwrite,
        CancellationToken cancellationToken,
        string? primaryOutputPath = null,
        string? sourcePath = null) {
        try {
            PrepareAssetsOutput(document, assetsPath, overwrite, primaryOutputPath, sourcePath);

            IReadOnlyList<OfficeDocumentMaterializedAsset> materialized = document.WriteAssetsToDirectory(
                assetsPath,
                new OfficeDocumentAssetMaterializationOptions {
                    CreateDirectory = true,
                    Overwrite = overwrite,
                    ValidatePayloadHash = true
                },
                cancellationToken);
            OfficeDocumentMaterializedAsset? unexpectedSkip = materialized.FirstOrDefault(result =>
                !result.Written &&
                result.Asset.PayloadBytes != null &&
                result.Asset.PayloadBytes.Length > 0 &&
                (string.IsNullOrWhiteSpace(result.Asset.PayloadHash) || result.Asset.PayloadHashMatches(out _)));
            if (unexpectedSkip != null) {
                throw new ReaderToolOutputException(
                    "Asset output could not be committed: " +
                    Path.GetFullPath(unexpectedSkip.Path ?? assetsPath));
            }
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new ReaderToolOutputException("Could not materialize document assets.", exception);
        }
    }

    internal static void PrepareAssetsOutput(
        OfficeDocumentReadResult document,
        string assetsPath,
        bool overwrite,
        string? primaryOutputPath = null,
        string? sourcePath = null) {
        try {
            if (File.Exists(assetsPath)) {
                throw new ReaderToolOutputException(
                    "Asset output path must be a directory: " + Path.GetFullPath(assetsPath));
            }
            if (!string.IsNullOrWhiteSpace(primaryOutputPath) &&
                primaryOutputPath != "-" &&
                IsSameOrChildPath(primaryOutputPath!, assetsPath)) {
                throw new ReaderToolOutputException(
                    "Asset directory cannot be the primary output path or one of its descendants.");
            }
            if (!string.IsNullOrWhiteSpace(sourcePath) &&
                IsSameOrChildPath(sourcePath!, assetsPath)) {
                throw new ReaderToolOutputException(
                    "Asset directory cannot be the input file path or one of its descendants.");
            }

            IReadOnlyList<string> outputPaths = GetMaterializableAssetPaths(document, assetsPath);
            for (int index = 0; index < outputPaths.Count; index++) {
                string outputPath = outputPaths[index];
                if (!string.IsNullOrWhiteSpace(primaryOutputPath) &&
                    primaryOutputPath != "-" &&
                    FilePathsConflict(primaryOutputPath!, outputPath)) {
                    throw new ReaderToolOutputException(
                        "Primary output and asset output paths conflict: " + Path.GetFullPath(outputPath));
                }
                if (!string.IsNullOrWhiteSpace(sourcePath) &&
                    FilePathsConflict(sourcePath!, outputPath)) {
                    throw new ReaderToolOutputException(
                        "Asset output conflicts with the input file: " + Path.GetFullPath(outputPath));
                }

                for (int previousIndex = 0; previousIndex < index; previousIndex++) {
                    if (OfficeImoToolPathSafety.PathsEqual(outputPaths[previousIndex], outputPath)) {
                        throw new ReaderToolOutputException(
                            "Multiple assets target the same output file: " + Path.GetFullPath(outputPath));
                    }
                }

                if (!overwrite && File.Exists(outputPath)) {
                    throw new ReaderToolOutputException(
                        "Asset output already exists. Use --force to replace it: " + Path.GetFullPath(outputPath));
                }
            }
            Directory.CreateDirectory(assetsPath);
        } catch (ReaderToolOutputException) {
            throw;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new ReaderToolOutputException("Could not prepare the asset output directory.", exception);
        }
    }

    private static IReadOnlyList<string> GetMaterializableAssetPaths(
        OfficeDocumentReadResult document,
        string assetsPath) {
        var outputPaths = new List<string>();
        foreach (OfficeDocumentAsset asset in document.Assets) {
            if (asset.PayloadBytes == null || asset.PayloadBytes.Length == 0) {
                continue;
            }
            if (!string.IsNullOrWhiteSpace(asset.PayloadHash) && !asset.PayloadHashMatches(out _)) {
                continue;
            }

            string fileName = string.IsNullOrWhiteSpace(asset.FileName)
                ? OfficeDocumentAssetNaming.BuildFileName(asset.Id, asset.Extension)
                : Path.GetFileName(asset.FileName!);
            if (string.IsNullOrWhiteSpace(fileName)) {
                fileName = OfficeDocumentAssetNaming.BuildFileName(asset.Id, asset.Extension);
            }
            outputPaths.Add(Path.Combine(assetsPath, fileName));
        }
        return outputPaths;
    }

    private static bool FilePathsConflict(string firstPath, string secondPath) =>
        IsSameOrChildPath(firstPath, secondPath) || IsSameOrChildPath(secondPath, firstPath);

    private static bool IsSameOrChildPath(string parentPath, string candidatePath) {
        string resolvedParent = OfficeImoToolPathSafety.ResolveExistingLinks(parentPath);
        string resolvedCandidate = OfficeImoToolPathSafety.ResolveExistingLinks(candidatePath);
        return OfficeImoToolPathSafety.IsSameOrChildPath(resolvedParent, resolvedCandidate);
    }

    internal static async Task WriteFileAsync(
        string path,
        string content,
        bool overwrite,
        CancellationToken cancellationToken) {
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
            File.Move(temporaryPath, fullPath, overwrite);
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
