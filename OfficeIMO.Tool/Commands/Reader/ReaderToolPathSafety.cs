namespace OfficeIMO.Tool.Commands.Reader;

internal static class ReaderToolPathSafety {
    internal static void EnsureDistinctFile(string inputPath, string outputPath) {
        try {
            if (!OfficeImoToolPathSafety.PathsEqual(inputPath, outputPath)) return;
            throw new ReaderToolOutputException("Output file must be different from the input file.");
        } catch (ReaderToolOutputException) {
            throw;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new ReaderToolOutputException("Could not resolve input and output paths.", exception);
        }
    }

    internal static void EnsureOutsideInput(string inputPath, params string?[] candidatePaths) {
        try {
            string resolvedInput = OfficeImoToolPathSafety.ResolveExistingLinks(inputPath);
            foreach (string? candidatePath in candidatePaths) {
                if (string.IsNullOrWhiteSpace(candidatePath)) continue;
                string resolvedCandidate = OfficeImoToolPathSafety.ResolveExistingLinks(candidatePath!);
                if (OfficeImoToolPathSafety.IsSameOrChildPath(resolvedInput, resolvedCandidate)) {
                    throw new ReaderToolOutputException(
                        "Folder output and asset directories must be outside the input folder.");
                }
            }
        } catch (ReaderToolOutputException) {
            throw;
        } catch (Exception exception) when (exception is IOException or UnauthorizedAccessException) {
            throw new ReaderToolOutputException("Could not resolve input and output paths.", exception);
        }
    }
}
