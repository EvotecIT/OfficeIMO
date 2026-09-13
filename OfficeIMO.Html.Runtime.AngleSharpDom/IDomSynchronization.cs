namespace AngleSharp.Dom;

/// <summary>Optional synchronization shared by parser mutations and a script host.</summary>
public interface IDomSynchronization {
    /// <summary>Gets the monitor used for synchronous DOM operations; never hold it across an await.</summary>
    object SyncRoot { get; }
}
