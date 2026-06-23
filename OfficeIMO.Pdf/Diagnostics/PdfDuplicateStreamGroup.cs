namespace OfficeIMO.Pdf;

/// <summary>Duplicate stream candidate group discovered by byte length and hash.</summary>
public sealed class PdfDuplicateStreamGroup {
    internal PdfDuplicateStreamGroup(string hash, long length, IReadOnlyList<int> objectNumbers) {
        Hash = hash;
        Length = length;
        ObjectNumbers = objectNumbers;
        DuplicateCount = Math.Max(0, objectNumbers.Count - 1);
        EstimatedSavingsBytes = Length * DuplicateCount;
    }

    /// <summary>Shared stream hash.</summary>
    public string Hash { get; }

    /// <summary>Length of each matching stream.</summary>
    public long Length { get; }

    /// <summary>Object numbers that share the same length and hash.</summary>
    public IReadOnlyList<int> ObjectNumbers { get; }

    /// <summary>Number of duplicate copies after keeping one representative stream.</summary>
    public int DuplicateCount { get; }

    /// <summary>Conservative byte saving estimate if duplicates could be shared.</summary>
    public long EstimatedSavingsBytes { get; }
}
