namespace OfficeIMO.Email;

/// <summary>Kind of privacy-safe semantic difference.</summary>
public enum EmailSemanticDifferenceKind {
    /// <summary>The source entry has no destination counterpart.</summary>
    MissingFromDestination = 0,
    /// <summary>The destination contains an entry absent from the source.</summary>
    UnexpectedInDestination = 1,
    /// <summary>The entry exists on both sides but its canonical value differs.</summary>
    Changed = 2
}

/// <summary>Describes one difference without exposing message values or content digests.</summary>
public sealed class EmailSemanticDifference {
    internal EmailSemanticDifference(string path, EmailSemanticDifferenceKind kind,
        long? sourceLength, long? destinationLength) {
        Path = path;
        Kind = kind;
        SourceLength = sourceLength;
        DestinationLength = destinationLength;
    }

    /// <summary>Stable canonical path of the differing semantic component.</summary>
    public string Path { get; }

    /// <summary>Difference classification.</summary>
    public EmailSemanticDifferenceKind Kind { get; }

    /// <summary>Source byte or character length when the component supplies one.</summary>
    public long? SourceLength { get; }

    /// <summary>Destination byte or character length when the component supplies one.</summary>
    public long? DestinationLength { get; }
}
