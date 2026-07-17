using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>One mismatched or unreadable item found during destination verification.</summary>
public sealed class EmailStorePstVerificationIssue {
    internal EmailStorePstVerificationIssue(string sourceItemId, string destinationItemId,
        bool isAssociated, string code, IReadOnlyList<EmailSemanticDifference> differences) {
        SourceItemId = sourceItemId;
        DestinationItemId = destinationItemId;
        IsAssociated = isAssociated;
        Code = code;
        Differences = differences;
    }

    /// <summary>Stable source-session item identifier.</summary>
    public string SourceItemId { get; }

    /// <summary>Stable destination PST item identifier.</summary>
    public string DestinationItemId { get; }

    /// <summary>Whether the item is folder-associated information.</summary>
    public bool IsAssociated { get; }

    /// <summary>Stable issue code.</summary>
    public string Code { get; }

    /// <summary>Value-free semantic differences. Empty when an item could not be read.</summary>
    public IReadOnlyList<EmailSemanticDifference> Differences { get; }
}
