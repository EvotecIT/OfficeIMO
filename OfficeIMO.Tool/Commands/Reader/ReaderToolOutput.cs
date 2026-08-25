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

        var primaryOutputPaths = new string[paths.Count];
        var assetDirectories = new string?[paths.Count];
        var plannedFilePaths = new List<string>();
        var plannedDirectoryPaths = new List<string>();
        for (int index = 0; index < paths.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            string relativePath = Path.GetRelativePath(sourceRoot, paths[index]);
            string suffix = format == ReaderToolOutputFormat.Json ? ".reader.json" : ".md";
            string outputPath = Path.Combine(outputRoot, relativePath + suffix);
            ReaderToolPathSafety.EnsureOutsideInput(sourceRoot, outputPath);
            primaryOutputPaths[index] = outputPath;
            AddPlannedFilePath(plannedFilePaths, outputPath);

            if (!string.IsNullOrWhiteSpace(assetsRoot) && documents[index].Assets.Count > 0) {
                string assetDirectory = Path.Combine(assetsRoot!, relativePath + ".assets");
                ReaderToolPathSafety.EnsureOutsideInput(sourceRoot, assetDirectory);
                assetDirectories[index] = assetDirectory;
                plannedDirectoryPaths.Add(assetDirectory);
                foreach (string assetPath in GetMaterializableAssetPaths(documents[index], assetDirectory)) {
                    AddPlannedFilePath(plannedFilePaths, assetPath);
                }
            }
        }

        foreach (string directoryPath in plannedDirectoryPaths) {
            foreach (string filePath in plannedFilePaths) {
                if (IsSameOrChildPath(filePath, directoryPath)) {
                    throw new ReaderToolOutputException(
                        "A planned output directory conflicts with a planned output file: " +
                        Path.GetFullPath(directoryPath));
                }
            }
        }

        for (int index = 0; index < paths.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            string outputPath = primaryOutputPaths[index];
            if (Directory.Exists(outputPath)) {
                throw new ReaderToolOutputException(
                    "Primary output path is a directory: " + Path.GetFullPath(outputPath));
            }
            ProbeWritableFilePath(outputPath);

            string? assetDirectory = assetDirectories[index];
            if (assetDirectory != null) {
                PrepareAssetsOutput(
                    documents[index],
                    assetDirectory,
                    overwrite: true,
                    outputPath,
                    paths[index]);
            }
        }

        for (int index = 0; index < paths.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            string outputPath = primaryOutputPaths[index];

            await WriteFileAsync(
                    outputPath,
                    FormatDocument(documents[index], format),
                    overwrite: true,
                    cancellationToken)
                .ConfigureAwait(false);

            string? assetDirectory = assetDirectories[index];
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
                    if (OutputPathsEquivalent(outputPaths[previousIndex], outputPath)) {
                        throw new ReaderToolOutputException(
                            "Multiple assets target the same output file: " + Path.GetFullPath(outputPath));
                    }
                }

                if (Directory.Exists(outputPath)) {
                    throw new ReaderToolOutputException(
                        "Asset output path is a directory: " + Path.GetFullPath(outputPath));
                }
                if (!overwrite && File.Exists(outputPath)) {
                    throw new ReaderToolOutputException(
                        "Asset output already exists. Use --force to replace it: " + Path.GetFullPath(outputPath));
                }
            }
            Directory.CreateDirectory(assetsPath);
            foreach (string outputPath in outputPaths) {
                ProbeWritableFilePath(outputPath);
            }
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

    private static void AddPlannedFilePath(ICollection<string> plannedPaths, string candidatePath) {
        foreach (string plannedPath in plannedPaths) {
            if (FilePathsConflict(plannedPath, candidatePath)) {
                throw new ReaderToolOutputException(
                    "Multiple outputs target conflicting file paths: " + Path.GetFullPath(candidatePath));
            }
        }
        plannedPaths.Add(candidatePath);
    }

    private static void ProbeWritableFilePath(string path) {
        try {
            string fullPath = Path.GetFullPath(path);
            if (!File.Exists(fullPath) && !Directory.Exists(fullPath)) {
                var destination = new FileInfo(fullPath);
                if (destination.LinkTarget != null) {
                    throw new ReaderToolOutputException(
                        "Output path is a dangling symbolic link: " + fullPath);
                }
            }

            string? directory = Path.GetDirectoryName(fullPath);
            if (string.IsNullOrWhiteSpace(directory)) {
                throw new ReaderToolOutputException("Output path must include a directory: " + fullPath);
            }

            Directory.CreateDirectory(directory);
            string intendedName = Path.GetFileName(fullPath);
            ValidateDestinationFileName(intendedName, fullPath);
            string probePath = File.Exists(fullPath)
                ? Path.Combine(directory, ".officeimo-write-probe-" + Guid.NewGuid().ToString("N"))
                : fullPath;
            using (new FileStream(
                       probePath,
                       FileMode.CreateNew,
                       FileAccess.Write,
                       FileShare.Delete,
                       bufferSize: 1,
                       FileOptions.DeleteOnClose)) {
            }
        } catch (ReaderToolOutputException) {
            throw;
        } catch (Exception exception) when (
            exception is IOException or UnauthorizedAccessException or ArgumentException or NotSupportedException) {
            throw new ReaderToolOutputException("Could not prepare output path '" + path + "'.", exception);
        }
    }

    private static void ValidateDestinationFileName(string fileName, string fullPath) {
        if (!OperatingSystem.IsWindows()) return;

        string windowsName = fileName.TrimEnd(' ', '.');
        string stem = windowsName.Split('.')[0];
        bool reserved =
            stem.Equals("CON", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("PRN", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("AUX", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("NUL", StringComparison.OrdinalIgnoreCase) ||
            (stem.Length == 4 &&
             (stem.StartsWith("COM", StringComparison.OrdinalIgnoreCase) ||
              stem.StartsWith("LPT", StringComparison.OrdinalIgnoreCase)) &&
             stem[3] is >= '1' and <= '9');
        if (windowsName.Length != fileName.Length || reserved) {
            throw new ReaderToolOutputException(
                "Output filename is not portable on Windows: " + fullPath);
        }
    }

    private static bool OutputPathsEquivalent(string firstPath, string secondPath) =>
        OfficeImoToolPathSafety.PathsEqual(firstPath, secondPath);

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
