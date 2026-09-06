namespace OfficeIMO.Workflows;

/// <summary>A user-selected provider destination with verified direct writes and durable recovery.</summary>
/// <remarks>Only Replace is supported. A provider cannot guarantee atomic replacement, exclusive creation, or rollback.</remarks>
public sealed class OfficeWorkflowStreamOutput {
    /// <summary>Creates a destination. The host must explain and obtain consent for direct provider writes.</summary>
    public OfficeWorkflowStreamOutput(string name, Func<CancellationToken, Task<Stream>> openRead,
        Func<CancellationToken, Task<Stream>> openWrite, OfficeWorkflowOutputRecoveryStore recoveryStore) {
        _ = new OfficeWorkflowStreamInput(name, openRead);
        OfficeWorkflowOutputRecoveryStore.ValidateExtension(name);
        Name = name;
        OpenRead = openRead;
        OpenWrite = openWrite ?? throw new ArgumentNullException(nameof(openWrite));
        RecoveryStore = recoveryStore ?? throw new ArgumentNullException(nameof(recoveryStore));
    }
    /// <summary>Gets the selected filename, including its output format extension.</summary>
    public string Name { get; }
    /// <summary>Gets the factory used to verify committed contents after closing the write stream.</summary>
    public Func<CancellationToken, Task<Stream>> OpenRead { get; }
    /// <summary>Gets the factory that opens a destructive write stream. The runner closes every returned stream.</summary>
    public Func<CancellationToken, Task<Stream>> OpenWrite { get; }
    /// <summary>Gets the durable store used before any destination write begins.</summary>
    public OfficeWorkflowOutputRecoveryStore RecoveryStore { get; }
}
