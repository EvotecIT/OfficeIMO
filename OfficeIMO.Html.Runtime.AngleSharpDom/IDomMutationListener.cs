namespace AngleSharp.Dom;

/// <summary>Optional provider hook for synchronous internal invalidation without scheduling a web observer.</summary>
public interface IDomMutationListener {
    /// <summary>Receives a mutation record at enqueue. Implementations must not mutate the DOM or execute script.</summary>
    void OnMutation(IDocument document, IMutationRecord record);
}
