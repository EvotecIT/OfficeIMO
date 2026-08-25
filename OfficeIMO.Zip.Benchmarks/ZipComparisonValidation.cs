namespace OfficeIMO.Zip.Benchmarks;

internal static class ZipComparisonValidation {
    internal static IReadOnlyList<ZipComparisonReport> ValidateAll() =>
        ZipComparisonCorpus.Scales.Select(Validate).ToArray();

    internal static ZipComparisonReport Validate(ZipBenchmarkScale scale) {
        byte[] input = ZipComparisonCorpus.CreateArchive(scale);
        return Validate(scale, input);
    }

    internal static ZipComparisonReport Validate(ZipBenchmarkScale scale, byte[] input) {
        ZipTraversalResult office = ZipComparisonWorkflows.TraverseOffice(input);
        IReadOnlyList<ZipProjectionDescriptor> platform = ZipComparisonWorkflows.TraversePlatform(input);
        if (office.Warnings.Count != 0) {
            throw new InvalidOperationException($"{scale.Name} OfficeIMO traversal produced unexpected warnings.");
        }
        if (office.EntriesVisited != scale.FileCount + scale.DirectoryCount) {
            throw new InvalidOperationException(
                $"{scale.Name} visited {office.EntriesVisited} entries; expected {scale.FileCount + scale.DirectoryCount}.");
        }

        ZipComparisonObservation officeObservation = ZipComparisonWorkflows.Observe(office.Entries);
        ZipComparisonObservation platformObservation = ZipComparisonWorkflows.Observe(platform);
        if (officeObservation != platformObservation) {
            throw new InvalidOperationException(
                $"{scale.Name} traversal projections differ: OfficeIMO={officeObservation}, platform={platformObservation}.");
        }
        if (officeObservation.EntryCount != scale.FileCount) {
            throw new InvalidOperationException(
                $"{scale.Name} returned {officeObservation.EntryCount} files; expected {scale.FileCount}.");
        }

        return new ZipComparisonReport(
            scale.Name,
            input.LongLength,
            office.EntriesVisited,
            officeObservation.EntryCount,
            officeObservation.TotalUncompressedBytes,
            officeObservation.StructuralFingerprint);
    }
}

internal sealed record ZipComparisonReport(
    string Scale,
    long InputBytes,
    int EntriesVisited,
    int EntryCount,
    long TotalUncompressedBytes,
    string StructuralFingerprint);
