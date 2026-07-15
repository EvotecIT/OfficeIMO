namespace OfficeIMO.Pdf;

/// <summary>Bounds the immutable raw syntax projection returned by the reader.</summary>
public sealed class PdfRawStructureOptions {
    private int _maxObjects = 10_000;
    private int _maxDepth = 12;
    private int _maxCollectionItems = 1_000;
    private int _maxTextLength = 4_096;

    /// <summary>Maximum indirect objects projected.</summary>
    public int MaxObjects { get => _maxObjects; set => _maxObjects = Positive(value, nameof(value)); }
    /// <summary>Maximum nested array, dictionary, and stream-dictionary depth.</summary>
    public int MaxDepth { get => _maxDepth; set => _maxDepth = Positive(value, nameof(value)); }
    /// <summary>Maximum entries projected from one array or dictionary.</summary>
    public int MaxCollectionItems { get => _maxCollectionItems; set => _maxCollectionItems = Positive(value, nameof(value)); }
    /// <summary>Maximum characters exposed from one PDF string or trailer preview.</summary>
    public int MaxTextLength { get => _maxTextLength; set => _maxTextLength = Positive(value, nameof(value)); }

    private static int Positive(int value, string parameterName) {
        if (value <= 0) throw new ArgumentOutOfRangeException(parameterName, value, "Raw structure limits must be positive.");
        return value;
    }
}
