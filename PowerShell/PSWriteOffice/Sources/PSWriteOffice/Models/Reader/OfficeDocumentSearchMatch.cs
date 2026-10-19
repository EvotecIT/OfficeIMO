using System;
using System.Collections.Generic;

namespace PSWriteOffice.Models.Reader;

/// <summary>PowerShell-friendly occurrence returned by a path-based document search.</summary>
public sealed class OfficeDocumentSearchMatch
{
    internal OfficeDocumentSearchMatch(
        string path,
        string documentType,
        string query,
        string? location,
        string text,
        string match,
        int startIndex,
        int length,
        IReadOnlyList<int> pages,
        bool documentLimitReached,
        bool sourceLimitReached,
        bool searchLimitReached)
    {
        Path = path;
        DocumentType = documentType;
        Query = query;
        Location = location;
        Text = text;
        Match = match;
        StartIndex = startIndex;
        Length = length;
        Pages = pages;
        DocumentLimitReached = documentLimitReached;
        SourceLimitReached = sourceLimitReached;
        SearchLimitReached = searchLimitReached;
    }

    /// <summary>Source file path.</summary>
    public string Path { get; }

    /// <summary>Detected document type.</summary>
    public string DocumentType { get; }

    /// <summary>Original search query.</summary>
    public string Query { get; }

    /// <summary>Logical location inside the source document.</summary>
    public string? Location { get; }

    /// <summary>Normalized source block containing the match.</summary>
    public string Text { get; }

    /// <summary>Exact matching text.</summary>
    public string Match { get; }

    /// <summary>Zero-based match offset within <see cref="Text"/>.</summary>
    public int StartIndex { get; }

    /// <summary>Match length.</summary>
    public int Length { get; }

    /// <summary>One-based physical pages containing the occurrence, when available.</summary>
    public IReadOnlyList<int> Pages { get; }

    /// <summary>True when the mixed-document scan stopped at its configured document ceiling.</summary>
    public bool DocumentLimitReached { get; }

    /// <summary>True when document ingestion reported a configured source limit.</summary>
    public bool SourceLimitReached { get; }

    /// <summary>True when the per-document search result ceiling was reached.</summary>
    public bool SearchLimitReached { get; }
}
