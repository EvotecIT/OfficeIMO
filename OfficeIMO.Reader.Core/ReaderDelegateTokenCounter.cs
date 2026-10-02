namespace OfficeIMO.Reader;

/// <summary>Optional token-counter extension that counts a prefix and source range without allocating a substring.</summary>
/// <remarks>The result must equal CountTokens(prefix + source.Substring(start, length)).</remarks>
public interface IReaderRangeTokenCounter : IReaderTokenCounter {
    /// <summary>Counts the exact concatenated output projection represented by a prefix and source range.</summary>
    int CountTokens(string prefix, string source, int start, int length);
}

/// <summary>Connects an application-owned tokenizer to Reader without adding a tokenizer package to its runtime graph.</summary>
/// <remarks>Delegates must be deterministic, thread-safe and return non-negative counts. Include the tokenizer model
/// or vocabulary version in the identifier so chunking evidence identifies the actual counting contract.</remarks>
public sealed class ReaderDelegateTokenCounter : IReaderRangeTokenCounter {
    private readonly Func<string, int> _count;
    private readonly Func<string, string, int, int, int>? _countRange;
    /// <summary>Creates a tokenizer bridge with an optional allocation-free range callback.</summary>
    public ReaderDelegateTokenCounter(string id, Func<string, int> count,
        Func<string, string, int, int, int>? countRange = null) {
        if (string.IsNullOrWhiteSpace(id)) throw new ArgumentException("A tokenizer identifier is required.", nameof(id));
        Id = id.Trim(); _count = count ?? throw new ArgumentNullException(nameof(count)); _countRange = countRange;
    }
    /// <inheritdoc />
    public string Id { get; }
    /// <inheritdoc />
    public int CountTokens(string text) => Validate(_count(text ?? string.Empty));
    /// <inheritdoc />
    public int CountTokens(string prefix, string source, int start, int length) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        if (start < 0 || length < 0 || start > source.Length - length) throw new ArgumentOutOfRangeException(nameof(start));
        return Validate(_countRange == null ? _count((prefix ?? string.Empty) + source.Substring(start, length))
            : _countRange(prefix ?? string.Empty, source, start, length));
    }
    private int Validate(int count) => count < 0
        ? throw new InvalidOperationException($"Token counter '{Id}' returned a negative count.") : count;
}
