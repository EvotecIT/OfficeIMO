namespace OfficeIMO.Epub.Benchmarks.Comparisons;

internal static class EpubComparisonValidation {
    internal static IReadOnlyList<EpubComparisonReport> ValidateAll() =>
        EpubComparisonCorpus.Scales.Select(scale => Validate(scale.Name)).ToArray();

    internal static EpubComparisonReport Validate(string scaleName) {
        EpubComparisonScale scale = EpubComparisonCorpus.Get(scaleName);
        byte[] package = EpubComparisonCorpus.CreatePackage(scale);
        EpubReadEvidence officeIMO = EpubComparisonWorkflows.InspectOfficeIMO(package);
        EpubReadEvidence versOne = EpubComparisonWorkflows.InspectVersOne(package);

        RequireEqual("title", officeIMO.Title, versOne.Title);
        RequireEqual("creator", officeIMO.Creator, versOne.Creator);
        RequireEqual("language", officeIMO.Language, versOne.Language);
        RequireEqual("chapter count", officeIMO.ChapterCount, versOne.ChapterCount);
        RequireEqual("content characters", officeIMO.ContentCharacters, versOne.ContentCharacters);
        RequireEqual("plain-text characters", officeIMO.TextCharacters, versOne.TextCharacters);
        RequireEqual("first path", officeIMO.FirstPath, versOne.FirstPath);
        RequireEqual("last path", officeIMO.LastPath, versOne.LastPath);
        RequireEqual("path hash", officeIMO.PathHash, versOne.PathHash);
        RequireEqual("content hash", officeIMO.ContentHash, versOne.ContentHash);
        RequireEqual("plain-text hash", officeIMO.TextHash, versOne.TextHash);
        RequireEqual("expected chapter count", scale.Chapters, officeIMO.ChapterCount);
        RequireEqual("expected title", EpubComparisonCorpus.Title, officeIMO.Title);
        RequireEqual("expected creator", EpubComparisonCorpus.Creator, officeIMO.Creator);
        RequireEqual("expected language", EpubComparisonCorpus.Language, officeIMO.Language);
        RequireEqual("first chapter path", EpubComparisonCorpus.ChapterPath(0), officeIMO.FirstPath);
        RequireEqual("last chapter path", EpubComparisonCorpus.ChapterPath(scale.Chapters - 1), officeIMO.LastPath);

        return new EpubComparisonReport(scale.Name, package.LongLength, officeIMO, versOne);
    }

    private static void RequireEqual<T>(string contract, T expected, T actual) {
        if (!EqualityComparer<T>.Default.Equals(expected, actual)) {
            throw new InvalidOperationException(
                $"EPUB {contract} differs. Expected '{expected}'; actual '{actual}'.");
        }
    }
}

internal sealed record EpubReadEvidence(
    string Implementation,
    long InputBytes,
    string? Title,
    string? Creator,
    string? Language,
    int ChapterCount,
    long ContentCharacters,
    long TextCharacters,
    string? FirstPath,
    string? LastPath,
    string PathHash,
    string ContentHash,
    string TextHash);

internal sealed record EpubComparisonReport(
    string Scale,
    long InputBytes,
    EpubReadEvidence OfficeIMO,
    EpubReadEvidence VersOne);
