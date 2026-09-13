namespace AngleSharp.Browser;

/// <summary>Optional event-loop hook for queuing DOM mutation notifications in a shared script microtask queue.</summary>
public interface IMutationMicrotaskScheduler {
    /// <summary>Queues one notification without running it synchronously.</summary>
    void EnqueueMutationMicrotask(System.Action notification);
}
