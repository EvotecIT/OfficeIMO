namespace OfficeIMO.Entities;

/// <summary>Controls how competing or intersecting readings are returned.</summary>
public enum EntityOverlapPolicy {
    /// <summary>Returns the best readings with alternative and contained readings attached.</summary>
    BestOnly,
    /// <summary>Returns each reading that reaches the confidence threshold.</summary>
    All,
}
