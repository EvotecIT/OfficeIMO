using System;

namespace OfficeIMO.Reader;

/// <summary>Counts tokens for deterministic chunk budgeting.</summary>
/// <remarks>
/// Implementations used concurrently must be thread-safe. Counts must be deterministic and non-negative.
/// </remarks>
public interface IReaderTokenCounter {
    /// <summary>Stable counter identifier retained in chunking evidence.</summary>
    string Id { get; }

    /// <summary>Counts tokens in the supplied text.</summary>
    int CountTokens(string text);
}

/// <summary>Dependency-free token estimate using approximately four characters per token.</summary>
public sealed class ReaderHeuristicTokenCounter : IReaderTokenCounter {
    private ReaderHeuristicTokenCounter() {
    }

    /// <summary>Shared stateless instance.</summary>
    public static ReaderHeuristicTokenCounter Instance { get; } = new ReaderHeuristicTokenCounter();

    /// <inheritdoc />
    public string Id => "officeimo.reader.heuristic-v1";

    /// <inheritdoc />
    public int CountTokens(string text) {
        if (string.IsNullOrEmpty(text)) return 0;
        return Math.Max(1, (text.Length + 3) / 4);
    }
}
